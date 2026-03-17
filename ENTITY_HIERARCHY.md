# T24 Entity Hierarchy — Parent-Child Requirements

All entities measured and generated for export **must** have a clear parent chain:

```
ZONE (polyline on T24-ZONE)
  └── WALL / SURFACE (circle on T24-WALL)
        └── OPENING (circle on T24-WIN / T24-DOOR / T24-SKY)
```

## XDATA Structure by Entity Type

### ZONE (LWPOLYLINE on T24-ZONE layer)
| Index | Code | Field       | Example          |
|-------|------|-------------|------------------|
| 0     | 1000 | "ZONE"      | type marker      |
| 1     | 1000 | zone-id     | "Z-001"          |
| 2     | 1000 | zone-name   | "BOILER ROOM"    |
| 0     | 1040 | ceiling-ht  | 9.0              |
| 0     | 1070 | floor       | 1                |
| 3     | 1000 | condition   | "Conditioned"    |
| 4     | 1000 | occupancy   | "Residential"    |

**Parent:** None (top-level)

### WALL (CIRCLE on T24-WALL layer)
| Index | Code | Field     | Example          |
|-------|------|-----------|------------------|
| 0     | 1000 | "WALL"    | type marker      |
| 1     | 1000 | wall-id   | "W-001"          |
| 2     | 1000 | wall-type | "Exterior Wall"  |
| 3     | 1000 | zone-id   | "Z-001"          |
| 0     | 1040 | wall-ht   | 9.0              |
| 1     | 1040 | len-ft    | 15.2             |
| 2     | 1040 | area-sqft | 136.8            |
| 3     | 1040 | azimuth   | 0.0 (degrees)    |
| 4-5   | 1040 | p1 (x,y)  | edge start       |
| 6-7   | 1040 | p2 (x,y)  | edge end         |

**Parent:** zone-id (1000 index 3)

### OPENING (CIRCLE on T24-WIN / T24-DOOR / T24-SKY layer)
| Index | Code | Field     | Example          |
|-------|------|-----------|------------------|
| 0     | 1000 | "OPENING" | type marker      |
| 1     | 1000 | open-id   | "O-001"          |
| 2     | 1000 | type      | "Window" / "Door" / "Skylight" |
| 3     | 1000 | zone-id   | "Z-001"          |
| 4     | 1000 | wall-id   | "W-001"          |
| 0     | 1040 | area-sqft | 30.0             |
| 1     | 1040 | u-factor  | 0.0              |
| 2     | 1040 | shgc      | 0.0              |

**Parent:** wall-id (1000 index 4) → zone-id (1000 index 3)

For skylights: wall-id = `"{zone-id}-ROOF"` (auto-generated horizontal surface)

## Auto-Generated Horizontal Surfaces (in tz_to_excel.py)

These are NOT drawn in AutoCAD — they're created during Excel export:
- `{zone-id}-SLAB` — Slab on Grade
- `{zone-id}-FLOOR` — Raised Floor
- `{zone-id}-CEIL` — Ceiling
- `{zone-id}-ROOF` — Roof (skylights parent to this)

## Parenting Rules

1. **Walls** auto-detect their zone via `tz-nearest-edge` at placement time
2. **Openings** auto-detect zone via `tz-nearest-edge` AND wall via `tz-nearest-wall-id` at placement time
3. **Skylights** auto-parent to `{zone-id}-ROOF`
4. The JSON export includes `zone_id` and `wall_id` fields on openings when available
5. `tz_to_excel.py` uses stored IDs when present, falls back to geometric matching for old data

## Verification

Run `TZ-LISTDATA` in AutoCAD to see the full tree:
```
ZONE Z-001  "Room Name"  322.0 sqft  Fl 1
  WALL W-001  Exterior Wall  15.2' x 9.0'h  136.8 sqft  0°
    OPEN O-001  Window  30.0 sqft
  WALL W-002  Exterior Wall  21.3' x 9.0'h  191.7 sqft  270°
    OPEN O-002  Door  21.0 sqft
```

Openings marked `(no wall)` or listed under `UNASSIGNED` need to be re-placed or manually fixed.
