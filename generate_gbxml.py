"""
generate_gbxml.py
Reads building data from an Excel template and generates a gbXML file
for import into EnergyPro 9.

Usage:
  python generate_gbxml.py                          <- creates blank template
  python generate_gbxml.py "My Project.xlsx"        <- generates gbXML from filled template
  python generate_gbxml.py "My Project.xlsx" out.xml
"""

import sys
import os
import xml.etree.ElementTree as ET
from xml.dom import minidom
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Orientation name -> azimuth degrees (clockwise from North)
ORIENTATION_MAP = {
    "north": 0,   "n": 0,
    "northeast": 45, "ne": 45,
    "east": 90,   "e": 90,
    "southeast": 135, "se": 135,
    "south": 180, "s": 180,
    "southwest": 225, "sw": 225,
    "west": 270,  "w": 270,
    "northwest": 315, "nw": 315,
    "front": 180,  # default front = south; user can override in Project tab
    "back": 0,
    "left": 90,
    "right": 270,
}

SURFACE_TYPE_MAP = {
    "exterior wall": "ExteriorWall",
    "ext wall": "ExteriorWall",
    "wall": "ExteriorWall",
    "interior wall": "InteriorWall",
    "int wall": "InteriorWall",
    "demising wall": "InteriorWall",
    "roof": "Roof",
    "ceiling": "Ceiling",
    "interior ceiling": "InteriorCeiling",
    "exposed floor": "ExposedFloor",
    "slab": "SlabOnGrade",
    "slab on grade": "SlabOnGrade",
    "raised floor": "RaisedFloor",
    "underground wall": "UndergroundWall",
    "underground slab": "UndergroundSlab",
}

OPENING_TYPE_MAP = {
    "window": "FixedWindow",
    "fixed window": "FixedWindow",
    "operable window": "OperableWindow",
    "skylight": "FixedSkylight",
    "door": "NonSlidingDoor",
    "sliding door": "SlidingDoor",
    "operable door": "NonSlidingDoor",
}


# ---------------------------------------------------------------------------
# Excel template creator
# ---------------------------------------------------------------------------

def make_header(ws, cols, fill_color="1F4E79", font_color="FFFFFF"):
    hdr_fill = PatternFill("solid", fgColor=fill_color)
    hdr_font = Font(bold=True, color=font_color)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col_idx, (header, width) in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

def add_note_row(ws, text, col_count, fill="D9E1F2"):
    r = ws.max_row + 1
    ws.cell(row=r, column=1, value="# " + text).font = Font(italic=True, color="595959")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=col_count)

def create_template(path: str):
    wb = Workbook()

    # ── Tab 1: Project ──────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Project"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 40

    info = [
        ("Project Name",       "My Building"),
        ("Address",            "123 Main St, Los Angeles, CA"),
        ("Climate Zone",       "CZ6  (Los Angeles)"),
        ("Building Type",      "MultiFamily"),
        ("Front Orientation",  "South"),
        ("Standards Version",  "2022"),
        ("Notes",              ""),
    ]
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    val_fill = PatternFill("solid", fgColor="EBF3FB")
    for r, (label, value) in enumerate(info, start=2):
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = Font(bold=True)
        lc.fill = hdr_fill
        lc.font = Font(bold=True, color="FFFFFF")
        vc = ws.cell(row=r, column=2, value=value)
        vc.fill = val_fill

    # ── Tab 2: Zones ────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Zones")
    zone_cols = [
        ("Zone ID\n(unique, no spaces)", 20),
        ("Zone Name",                    28),
        ("Floor Area (sqft)",            18),
        ("Ceiling Height (ft)",          18),
        ("Type\n(Conditioned/Unconditioned)", 24),
    ]
    make_header(ws2, zone_cols)
    sample_zones = [
        ("Z-101", "Unit 101", 651,  8, "Conditioned"),
        ("Z-102", "Unit 102", 900,  8, "Conditioned"),
        ("Z-LOBBY", "Lobby",  400,  10, "Conditioned"),
    ]
    for r, row in enumerate(sample_zones, start=2):
        for c, val in enumerate(row, start=1):
            ws2.cell(row=r, column=c, value=val)
    add_note_row(ws2, "Zone ID must be unique and contain no spaces. Used to link walls to zones.", 5)

    # ── Tab 3: Walls ────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Walls")
    wall_cols = [
        ("Wall ID\n(unique, no spaces)", 20),
        ("Zone ID",                      16),
        ("Wall Name",                    24),
        ("Type\n(Exterior Wall / Interior Wall / Roof / Slab on Grade)", 30),
        ("Orientation\n(N/S/E/W or degrees)", 22),
        ("Gross Area (sqft)",            18),
        ("Construction",                 30),
        ("Adjacent Zone ID\n(interior walls only)", 24),
    ]
    make_header(ws3, wall_cols)
    sample_walls = [
        ("W-101-N",  "Z-101", "North Wall",  "Exterior Wall", "North",  133, "R-21 Wood Framed Wall", ""),
        ("W-101-S",  "Z-101", "South Wall",  "Exterior Wall", "South",  133, "R-21 Wood Framed Wall", ""),
        ("W-101-E",  "Z-101", "East Wall",   "Exterior Wall", "East",   100, "R-21 Wood Framed Wall", ""),
        ("W-101-DM", "Z-101", "Demising",    "Interior Wall", "",       220, "R-0 Wall",              "Z-102"),
        ("W-101-RF", "Z-101", "Roof",        "Roof",          "",       651, "R-38 Roof",             ""),
        ("W-101-SL", "Z-101", "Slab",        "Slab on Grade", "",       651, "Slab-on-Grade",         ""),
    ]
    for r, row in enumerate(sample_walls, start=2):
        for c, val in enumerate(row, start=1):
            ws3.cell(row=r, column=c, value=val)
    add_note_row(ws3, "Gross Area = total wall area INCLUDING windows/doors. Orientation required for exterior walls.", 8)

    # ── Tab 4: Openings ─────────────────────────────────────────────────────
    ws4 = wb.create_sheet("Openings")
    open_cols = [
        ("Opening ID\n(unique, no spaces)", 22),
        ("Wall ID",                         16),
        ("Opening Name",                    24),
        ("Type\n(Window/Door/Skylight)",     22),
        ("Area (sqft)",                      14),
        ("U-Factor",                         12),
        ("SHGC",                             10),
    ]
    make_header(ws4, open_cols)
    sample_openings = [
        ("O-101-N-1", "W-101-N", "North Window A", "Window", 27,  0.27, 0.18),
        ("O-101-N-2", "W-101-N", "North Window B", "Window", 27,  0.27, 0.18),
        ("O-101-S-1", "W-101-S", "South Window A", "Window", 48,  0.27, 0.18),
        ("O-101-E-1", "W-101-E", "East Door",      "Door",   24,  0.50, ""),
    ]
    for r, row in enumerate(sample_openings, start=2):
        for c, val in enumerate(row, start=1):
            ws4.cell(row=r, column=c, value=val)
    add_note_row(ws4, "U-Factor and SHGC required for windows; leave blank for doors.", 7)

    wb.save(path)
    print(f"Template created: {path}")


# ---------------------------------------------------------------------------
# gbXML generator
# ---------------------------------------------------------------------------

def cell_val(ws, row, col):
    v = ws.cell(row=row, column=col).value
    return v if v is not None else ""

def resolve_azimuth(orientation_str) -> float:
    """Convert orientation string or number to azimuth degrees."""
    if orientation_str == "" or orientation_str is None:
        return 0.0
    try:
        return float(orientation_str)
    except (ValueError, TypeError):
        key = str(orientation_str).strip().lower()
        return float(ORIENTATION_MAP.get(key, 0))

def resolve_surface_type(type_str) -> str:
    key = str(type_str).strip().lower()
    return SURFACE_TYPE_MAP.get(key, "ExteriorWall")

def resolve_opening_type(type_str) -> str:
    key = str(type_str).strip().lower()
    return OPENING_TYPE_MAP.get(key, "FixedWindow")

def tilt_for_surface(surface_type: str) -> float:
    """Returns tilt in degrees (0=horizontal, 90=vertical)."""
    if surface_type in ("Roof", "SlabOnGrade", "ExposedFloor", "RaisedFloor", "UndergroundSlab", "Ceiling", "InteriorCeiling"):
        return 0.0
    return 90.0  # walls

def generate_gbxml(xlsx_path: str, out_path: str):
    wb = openpyxl.load_workbook(xlsx_path)

    # -- Project info --
    ws_proj = wb["Project"]
    def proj(label_row):
        return str(ws_proj.cell(row=label_row, column=2).value or "")

    project_name    = proj(2)
    address         = proj(3)
    climate_zone    = proj(4)
    building_type   = proj(5) or "MultiFamily"
    front_orient    = proj(6) or "South"
    standards_ver   = proj(7) or "2022"

    # -- Zones --
    ws_zones = wb["Zones"]
    zones = []
    for row in ws_zones.iter_rows(min_row=2, values_only=True):
        if not row[0] or str(row[0]).startswith("#"): continue
        zone_id, name, area, height, ztype = row[0], row[1], row[2], row[3], row[4]
        if not zone_id: continue
        zones.append({
            "id":     str(zone_id).strip().replace(" ", "_"),
            "name":   str(name or zone_id),
            "area":   float(area or 0),
            "height": float(height or 9),
            "type":   str(ztype or "Conditioned"),
        })

    # -- Walls --
    ws_walls = wb["Walls"]
    walls = []
    for row in ws_walls.iter_rows(min_row=2, values_only=True):
        if not row[0] or str(row[0]).startswith("#"): continue
        wid, zid, name, wtype, orient, area, construction, adj_zone = row[:8]
        if not wid: continue
        stype = resolve_surface_type(wtype or "Exterior Wall")
        walls.append({
            "id":           str(wid).strip().replace(" ", "_"),
            "zone_id":      str(zid or "").strip().replace(" ", "_"),
            "name":         str(name or wid),
            "surface_type": stype,
            "azimuth":      resolve_azimuth(orient),
            "tilt":         tilt_for_surface(stype),
            "area":         float(area or 0),
            "construction": str(construction or ""),
            "adj_zone":     str(adj_zone or "").strip().replace(" ", "_"),
        })

    # -- Openings --
    ws_open = wb["Openings"]
    openings = []
    for row in ws_open.iter_rows(min_row=2, values_only=True):
        if not row[0] or str(row[0]).startswith("#"): continue
        oid, wall_id, name, otype, area, ufactor, shgc = row[:7]
        if not oid: continue
        openings.append({
            "id":       str(oid).strip().replace(" ", "_"),
            "wall_id":  str(wall_id or "").strip().replace(" ", "_"),
            "name":     str(name or oid),
            "type":     resolve_opening_type(otype or "Window"),
            "area":     float(area or 0),
            "ufactor":  float(ufactor) if ufactor not in ("", None) else None,
            "shgc":     float(shgc) if shgc not in ("", None) else None,
        })

    # Index openings by wall_id
    openings_by_wall = {}
    for o in openings:
        openings_by_wall.setdefault(o["wall_id"], []).append(o)

    # -- Build gbXML tree --
    root = ET.Element("gbXML", {
        "xmlns":           "http://www.gbxml.org/schema",
        "temperatureUnit": "F",
        "lengthUnit":      "Feet",
        "areaUnit":        "SquareFeet",
        "volumeUnit":      "CubicFeet",
        "version":         "6.01",
    })

    campus = ET.SubElement(root, "Campus", {"id": "campus-1"})
    ET.SubElement(campus, "Name").text = project_name
    ET.SubElement(campus, "Location").text = address

    total_area = sum(z["area"] for z in zones)
    building = ET.SubElement(campus, "Building", {
        "id":           "building-1",
        "buildingType": building_type,
    })
    ET.SubElement(building, "Name").text = project_name
    ET.SubElement(building, "Area").text = str(total_area)

    # Spaces
    for z in zones:
        volume = z["area"] * z["height"]
        space = ET.SubElement(building, "Space", {
            "id":        z["id"],
            "zoneIdRef": z["id"],
        })
        ET.SubElement(space, "Name").text    = z["name"]
        ET.SubElement(space, "Area").text    = str(z["area"])
        ET.SubElement(space, "Volume").text  = str(round(volume, 2))
        ET.SubElement(space, "CeilingHeight").text = str(z["height"])

    # Surfaces (at Campus level, referencing spaces)
    for w in walls:
        surf = ET.SubElement(campus, "Surface", {
            "id":          w["id"],
            "surfaceType": w["surface_type"],
        })
        ET.SubElement(surf, "Name").text    = w["name"]
        ET.SubElement(surf, "Area").text    = str(w["area"])
        ET.SubElement(surf, "Azimuth").text = str(w["azimuth"])
        ET.SubElement(surf, "Tilt").text    = str(w["tilt"])

        if w["construction"]:
            ET.SubElement(surf, "CADObjectId").text = w["construction"]

        # Adjacent spaces
        if w["zone_id"]:
            ET.SubElement(surf, "AdjacentSpaceId", {"spaceIdRef": w["zone_id"]})
        if w["adj_zone"]:
            ET.SubElement(surf, "AdjacentSpaceId", {"spaceIdRef": w["adj_zone"]})

        # Openings
        for o in openings_by_wall.get(w["id"], []):
            opening = ET.SubElement(surf, "Opening", {
                "id":          o["id"],
                "openingType": o["type"],
            })
            ET.SubElement(opening, "Name").text = o["name"]
            ET.SubElement(opening, "Area").text = str(o["area"])
            if o["ufactor"] is not None:
                ET.SubElement(opening, "U-value").text = str(o["ufactor"])
            if o["shgc"] is not None:
                ET.SubElement(opening, "SHGC").text = str(o["shgc"])

    # -- Write pretty XML --
    rough = ET.tostring(root, encoding="unicode")
    # Strip any auto-injected empty namespace that Python adds
    rough = rough.replace(' xmlns=""', '').replace(" xmlns=''", '')
    pretty = minidom.parseString(rough).toprettyxml(indent="  ")
    lines = pretty.split("\n")
    if lines[0].startswith("<?xml"):
        lines[0] = '<?xml version="1.0" encoding="utf-8"?>'
    output = "\n".join(lines)
    # Clean up any empty xmlns that Python might add to child elements
    output = output.replace(' xmlns=""', '').replace(" xmlns=''", '')

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(output)

    print(f"gbXML generated: {out_path}")
    print(f"  {len(zones)} zone(s), {len(walls)} surface(s), {len(openings)} opening(s)")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    args = sys.argv[1:]

    if not args:
        # Create template
        template_path = os.path.join(SCRIPT_DIR, "T24 Input Template.xlsx")
        create_template(template_path)
        print("\nFill in the template, then run:")
        print(f'  python generate_gbxml.py "T24 Input Template.xlsx"')
        return

    xlsx_path = args[0]
    if not os.path.isabs(xlsx_path):
        xlsx_path = os.path.join(SCRIPT_DIR, xlsx_path)

    if not os.path.exists(xlsx_path):
        print(f"File not found: {xlsx_path}")
        sys.exit(1)

    if len(args) >= 2:
        out_path = args[1]
    else:
        base = os.path.splitext(xlsx_path)[0]
        out_path = base + ".xml"

    generate_gbxml(xlsx_path, out_path)


if __name__ == "__main__":
    main()
