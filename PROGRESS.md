# T24 TakeOff Wizard — Session Progress

Last updated: 2026-02-27
Latest commit: `d54eebe` (pushed to github.com/ddilanchi/t24-streamline)

---

## What Works Right Now

- **TZ-ZONE** — Click room name text → auto-traces boundary → tags zone with XDATA
- **TZ-WIN / TZ-DOOR / TZ-SKY** — Click to place opening markers
- **TZ-EXPORT** — Writes `<drawing>_t24.json`
- **tz_to_excel.py** — Converts JSON → Excel
- **generate_gbxml.py** — Excel → gbXML for EnergyPro
- **viewer_server.py + viewer.html** — 3D preview at localhost:5000

---

## Recent Work (this session)

### 1. Door-collapse rewrite (`eec5dd8`)
Old `tz-clean-pline` used an arc-first approach: found each arc, searched ±8
vertices by index count for a door-width segment. Picked the wrong door when
multiple doors were nearby.

New approach is **door-first with geometric distance pairing**:
- Collect all arc vertex positions and door-width segments (30–44") before flattening
- For each door segment, find nearest unused arc by Euclidean distance (max 72")
- Remove shorter polygon path between arc index and door-segment start (the notch)
- `used-arcs` list prevents two doors competing for the same arc

### 2. Spike removal (`d54eebe`)
BOUNDARY sometimes creates a there-and-back spur to inner closed shapes
(room-number rectangles, block attribute borders). This left a V-spike in the
boundary polyline that door-collapse never touched because it has no arc.

**Detection criteria (all three must be true):**
- `dist(Vi, Vj) < 24"` — bracket points close together on same wall
- `path_len(Vi→Vj) >= 3 × direct` — the detour is much longer than going straight
- `path_len >= 36"` — minimum spike size

Runs in a loop until no more spikes remain. Fires as Step 0 in `tz-clean-pline`,
before arc/door detection.

**Triggered by:** MED PREP room — room number "114" has a rectangle box whose
corner overlapped the door-arc area. BOUNDARY bridged to it and created a spike.

### 3. Room-number-finder fix (`9e39c83`)
Two bugs caused the room-number auto-append to stop working after a few rooms:
- `txt-lyr` was not reset between loop iterations. Clicking a non-text entity
  carried the previous room's layer name into the next ssget search.
- `nentsel` (with DRAWORDER-front T24 entities) was returning our own T24-LABEL
  text as the zone name. Now guarded: any entity on a `T24-*` layer prints a
  warning and skips.

### 4. Layer thaw moved to session end (`9e39c83`)
Thaw calls moved out of the per-room loop to after the `while` loop. Any frozen
layers stay frozen for the entire TZ-ZONE session and only release at the end.
(Currently `froze` and `txt-lyrs-frozen` are always nil — nothing is being
frozen — but the structure is correct for future use.)

---

## Current `tz-clean-pline` Step Order

```
Step 0 — Spike removal (there-and-back detour detection)
Step 1 — Collect arc positions + door-width segments (30–44") from raw verts
Step 2 — Flatten all arcs to straight segments
Step 3 — Pair each door with nearest unused arc (max 72" distance)
Step 4 — Remove shorter polygon path between arc and door (the notch)
Step 5 — Rebuild polyline from remaining vertices
```

---

## Known Remaining Issues / Next Steps

- **Spike tuning** — thresholds (24", 3×, 36") may need adjustment depending on
  drawing scale. If a room looks wrong, adjust these in `tz-clean-pline` Step 0.
- **TZ-WALL command** — not yet implemented. Walls are currently auto-generated
  from polyline edges in `tz_to_excel.py`. The plan (in the plan file) calls for
  a manual click-based TZ-WALL command similar to TZ-WIN.
- **4 horizontal surfaces per zone** — plan calls for Slab on Grade, Interior
  Floor, Interior Ceiling, Roof per zone in `tz_to_excel.py` (currently 2).
- **XDATA "T" bug** — plan mentions switching remaining quoted XDATA pairs to
  `(cons ...)` form. May already be resolved; verify on export.

---

## Key File Locations

| File | Path |
|------|------|
| Primary LSP | `D:\Documents\T24 Streamline Project\T24_TakeOff.lsp` |
| Python scripts | `D:\Documents\T24 Streamline Project\` |
| Git remote | github.com/ddilanchi/t24-streamline |

---

## Critical Rules

- **NEVER freeze, thaw, or modify AutoCAD layers** unless explicitly asked.
  A-AREA-IDEN and similar AIA layers contain BOTH text labels AND room boundary
  polylines. Freezing them breaks `_-BOUNDARY`.
- Always check paren balance before committing the LSP.
- Commit and push after each working milestone.
