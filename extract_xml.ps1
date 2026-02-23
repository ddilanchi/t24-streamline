# extract_xml.ps1
# Extracts ComplianceXML from EnergyPro .bld files using a compiled C# helper
# that hooks AssemblyResolve before deserialization begins.

param([string]$BldFile = "")

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BldDir    = Join-Path $ScriptDir "BLD EXAMPLES"
$OutDir    = Join-Path $ScriptDir "extracted_xml"
$EP9Dir    = "C:\Program Files (x86)\EnergySoft\EnergyPro 9"

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

# --- Compile C# helper (all deserialization logic in native .NET) ---
$csCode = @"
using System;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Collections.Generic;

public static class BldExtractor
{
    private static bool _hooked = false;

    public static void LoadAssemblies(string dir)
    {
        // Hook BEFORE loading anything so the handler is in place
        if (!_hooked)
        {
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
            _hooked = true;
        }

        foreach (string dll in Directory.GetFiles(dir, "EnergySoft.*.dll"))
        {
            try { Assembly.LoadFrom(dll); } catch { }
        }
        foreach (string dll in Directory.GetFiles(dir, "DataTypes.dll"))
        {
            try { Assembly.LoadFrom(dll); } catch { }
        }
    }

    private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
    {
        var requested = new AssemblyName(args.Name);
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
        {
            if (string.Equals(asm.GetName().Name, requested.Name,
                              StringComparison.OrdinalIgnoreCase))
            {
                return asm;
            }
        }
        return null;
    }

    public static string ExtractXml(string bldPath)
    {
        using (var stream = File.OpenRead(bldPath))
        {
            var formatter = new BinaryFormatter();
            object obj = formatter.Deserialize(stream);

            // Try ComplianceXML property
            PropertyInfo prop = obj.GetType().GetProperty("ComplianceXML");
            if (prop != null)
            {
                object val = prop.GetValue(obj);
                if (val != null)
                {
                    string xml = val.ToString();
                    if (xml.Length > 100) return xml;
                }
            }

            // Try backing field directly
            FieldInfo field = obj.GetType().GetField("<ComplianceXML>k__BackingField",
                                  BindingFlags.NonPublic | BindingFlags.Instance);
            if (field != null)
            {
                object val = field.GetValue(obj);
                if (val != null)
                {
                    string xml = val.ToString();
                    if (xml.Length > 100) return xml;
                }
            }

            // List available members for diagnostics
            var members = new List<string>();
            foreach (var p in obj.GetType().GetProperties())
                members.Add("prop:" + p.Name);
            foreach (var f in obj.GetType().GetFields(BindingFlags.NonPublic | BindingFlags.Instance))
                members.Add("field:" + f.Name);
            return "NULL|" + string.Join(", ", members);
        }
    }
}
"@

try {
    Add-Type -TypeDefinition $csCode -Language CSharp
    Write-Host "C# helper compiled OK"
} catch {
    Write-Host "Compile error: $($_.Exception.Message)"
    exit 1
}

# Load DLLs (hooks AssemblyResolve first)
[BldExtractor]::LoadAssemblies($EP9Dir)
Write-Host "Assemblies loaded"

# Pick files
if ($BldFile -ne "") {
    $targets = @($BldFile)
} else {
    $targets = Get-ChildItem -Path $BldDir -Filter "*.bld" |
               Select-Object -ExpandProperty FullName
}

foreach ($path in $targets) {
    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "File: $(Split-Path -Leaf $path)"

    try {
        $result = [BldExtractor]::ExtractXml($path)

        if ($result.StartsWith("NULL|")) {
            Write-Host "  ComplianceXML not found. Members:"
            $result.Substring(5).Split(",") | ForEach-Object { Write-Host "    $($_.Trim())" }
        } else {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
            $outPath  = Join-Path $OutDir "$baseName.xml"
            [System.IO.File]::WriteAllText($outPath, $result, [System.Text.Encoding]::UTF8)
            Write-Host "  [OK] $($result.Length) chars -> $(Split-Path -Leaf $outPath)"
        }
    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "Done. Output: $OutDir"
