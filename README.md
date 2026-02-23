# T24 Streamline Project

Tooling to reduce manual data entry for California Title 24 (EnergyPro) compliance work.

## The Problem

T24 compliance requires manually measuring room areas, wall areas, and window dimensions from architectural plans, then re-entering all of it by hand into EnergyPro. This is slow and error-prone.

## The Solution

Fill in one Excel spreadsheet → run one command → import into EnergyPro 9.

## Quick Start

**1. Generate the input template (first time only):**
```
python generate_gbxml.py
```
This creates `T24 Input Template.xlsx`.

**2. Fill in the template** — 4 tabs:
| Tab | What to enter |
|-----|--------------|
| Project | Name, address, climate zone, building type |
| Zones | One row per zone/unit — floor area, ceiling height |
| Walls | One row per wall — type, orientation, gross area, construction |
| Openings | One row per window/door — area, U-factor, SHGC |

**3. Generate the gbXML:**
```
python generate_gbxml.py "My Project.xlsx"
```
Outputs `My Project.xml`.

**4. Import into EnergyPro 9:**
File → Import → Green Building XML → select the `.xml` file.

## File Structure

```
generate_gbxml.py       Main tool — reads Excel, writes gbXML
T24 Input Template.xlsx Blank input template (sample data pre-filled)
extract_xml.py          Utility — attempts to extract SDDXML from .bld files
extract_xml.ps1         PowerShell version of the .bld extractor
```

## Orientation Reference

| Label | Azimuth |
|-------|---------|
| North / N | 0° |
| East / E | 90° |
| South / S | 180° |
| West / W | 270° |

Diagonal directions (NE, SW, etc.) and numeric degrees are also accepted.

## Surface Types

`Exterior Wall`, `Interior Wall`, `Roof`, `Ceiling`, `Slab on Grade`, `Exposed Floor`, `Raised Floor`

## Opening Types

`Window`, `Fixed Window`, `Operable Window`, `Skylight`, `Door`

## Requirements

- Python 3.x
- openpyxl (`pip install openpyxl`)
- EnergyPro 9 (for import)

## Notes

- gbXML namespace used: `http://www.gbxml.org/schema` (required by EnergyPro 9)
- Units: square feet, feet
- gbXML version: 6.01
- EnergyPro 10 does not appear to support gbXML import
