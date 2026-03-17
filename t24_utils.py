"""t24_utils.py — Shared constants and helpers for T24 Streamline pipeline."""

# Orientation name -> azimuth degrees (clockwise from North)
ORIENTATION_MAP = {
    "north": 0,    "n": 0,
    "northeast": 45,  "ne": 45,
    "east": 90,    "e": 90,
    "southeast": 135, "se": 135,
    "south": 180,  "s": 180,
    "southwest": 225, "sw": 225,
    "west": 270,   "w": 270,
    "northwest": 315, "nw": 315,
    "front": 180,
    "back": 0,
    "left": 90,
    "right": 270,
}


def parse_azimuth(s):
    """Convert an orientation string or numeric string to a float azimuth.
    Returns 0.0 for None/empty/unrecognized values."""
    if s in (None, ""):
        return 0.0
    try:
        return float(s)
    except (ValueError, TypeError):
        return float(ORIENTATION_MAP.get(str(s).strip().lower(), 0))


def az_to_cardinal(az):
    """Return compass direction string for the azimuth."""
    az = az % 360
    if az < 22.5 or az >= 337.5: return 'North'
    if az < 67.5:  return 'NE'
    if az < 112.5: return 'East'
    if az < 157.5: return 'SE'
    if az < 202.5: return 'South'
    if az < 247.5: return 'SW'
    if az < 292.5: return 'West'
    return 'NW'
