# T24 Streamline Project

Tooling to reduce manual data entry for California Title 24 (EnergyPro) compliance work.

## The Problem

T24 compliance requires manually measuring room areas, wall areas, and window dimensions from architectural plans, then re-entering all of it by hand into EnergyPro. This is slow and error-prone.

## The Solution

Two workflows depending on whether you have AutoCAD drawings:

### Workflow A: AutoCAD → EnergyPro (recommended)

1. **Load** `T24_TakeOff.lsp` in AutoCAD (APPLOAD or drag-drop)
2. **TZ-ZONE** — Click room name text to auto-trace boundaries
3. **TZ-WALL** — Click polyline edges to place wall markers
4. **TZ-WIN / TZ-DOOR / TZ-SKY** — Click to place opening markers
5. **TZ-EXPORT** — Writes `<drawing>_t24.json`
6. **Run** `python tz_to_excel.py "<drawing>_t24.json"` → generates Excel + geometry sidecar
7. **Review** Excel, fill in Construction column
8. **Run** `python generate_gbxml.py "<takeoff>.xlsx" output.xml` → generates gbXML
9. **Import** into EnergyPro 9: File → Import → Green Building XML

### Workflow B: Manual Excel → EnergyPro

1. **Run** `python generate_gbxml.py` to create a blank `T24 Input Template.xlsx`
2. **Fill in** the 4 tabs (Project, Zones, Walls, Openings)
3. **Run** `python generate_gbxml.py "My Project.xlsx" output.xml`
4. **Import** into EnergyPro 9

### 3D Viewer (optional)

Run `python viewer_server.py` and open `http://localhost:5000` to preview the building geometry in 3D.

## File Structure

```
T24_TakeOff.lsp         AutoCAD LISP — tags zones, walls, openings in the drawing
tz_to_excel.py          Converts AutoCAD JSON export → Excel + geometry sidecar
generate_gbxml.py       Reads Excel spreadsheet → writes gbXML for EnergyPro
viewer.html             3D building geometry viewer (Three.js)
viewer_server.py        Local server for the 3D viewer
T24 Input Template.xlsx Blank input template (for manual workflow)
extract_xml.py          Utility — extracts SDDXML from .bld files
extract_xml.ps1         PowerShell version of the .bld extractor
```

## AutoCAD Commands

| Command | Description |
|---------|-------------|
| `TZ-ZONE` | Click room name text; auto-traces boundary, assigns zone ID |
| `TZ-WALL` | Click polyline edges; places wall markers with length/azimuth |
| `TZ-WIN` | Click window locations; prompts for area |
| `TZ-DOOR` | Click door location; prompts for area |
| `TZ-SKY` | Click inside zone for skylight; prompts for area, U-factor, SHGC |
| `TZ-EXPORT` | Writes JSON file with all tagged data |
| `TZ-LISTDATA` | Lists all zone data in the command line |
| `TZ-SHOWVERTS` | Debug: labels polyline vertices with interior angles |
| `TZ-RESET` | Deletes all T24 labels, markers, and wall markers |

## Data Schema

### JSON Export (`_t24.json`)

- **zones[]** — `{id, name, vertices, centroid, area_sqft, ceiling_ht_ft, floor, condition, occupancy}`
- **walls[]** — `{id, type, zone_id, height_ft, length_ft, area_sqft, azimuth, p1, p2}`
- **openings[]** — `{id, type, position, area_sqft, u_factor, shgc}`

### Excel Spreadsheet

| Tab | What to enter |
|-----|--------------|
| Project | Name, address, climate zone, building type |
| Zones | One row per zone/unit — floor area, ceiling height |
| Walls | One row per wall — type, orientation, gross area, construction |
| Openings | One row per window/door — area, U-factor, SHGC |

Each zone gets 4 horizontal surfaces automatically: Slab on Grade, Interior Floor, Interior Ceiling, Roof.

### XDATA Tags

All AutoCAD entities use the `T24TOW` application ID. XDATA is stored using `(cons)` pairs (not quoted pairs) to avoid the boolean `T` round-trip bug in some AutoCAD versions.

## AutoCAD Layers

| Layer | Color | Content |
|-------|-------|---------|
| T24-ZONE | Cyan (4) | Zone boundary polylines |
| T24-WALL | Green (3) | Wall marker circles |
| T24-WIN | Yellow (2) | Window marker circles |
| T24-DOOR | Red (1) | Door marker circles |
| T24-SKY | Blue (5) | Skylight marker circles |
| T24-LABEL | White (7) | Label text |

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
- AutoCAD (for Workflow A)
- EnergyPro 9 (for import)

## Notes

- gbXML namespace: `http://www.gbxml.org/schema` (required by EnergyPro 9)
- Units: square feet, feet
- gbXML version: 6.01
- EnergyPro 10 does not appear to support gbXML import
