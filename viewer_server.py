"""viewer_server.py – T24 3D live viewer/editor
Run: python viewer_server.py  →  http://localhost:5000
"""
import sys, os, time, subprocess, threading, webbrowser

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

try:
    from flask import Flask, jsonify, request, send_from_directory
    import openpyxl
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "flask", "openpyxl"], check=True)
    from flask import Flask, jsonify, request, send_from_directory
    import openpyxl

app = Flask(__name__)

ORIENT = {
    "north":0,"n":0,"ne":45,"northeast":45,"east":90,"e":90,
    "se":135,"southeast":135,"south":180,"s":180,"sw":225,"southwest":225,
    "west":270,"w":270,"nw":315,"northwest":315,"front":180,"back":0,
}

def azimuth(s):
    if s in (None, ""): return 0.0
    try: return float(s)
    except: return float(ORIENT.get(str(s).strip().lower(), 0))

def flt(v, default=0.0):
    try: return float(v) if v not in (None, "") else default
    except: return default

def find_xlsx():
    return sorted(f for f in os.listdir(SCRIPT_DIR)
                  if f.lower().endswith(".xlsx") and not f.startswith("~$"))

def read_data(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    p = wb["Project"]
    project = {
        "name": str(p.cell(2,2).value or ""),
        "address": str(p.cell(3,2).value or ""),
        "climate_zone": str(p.cell(4,2).value or ""),
        "building_type": str(p.cell(5,2).value or "MultiFamily"),
        "front_orientation": str(p.cell(6,2).value or "South"),
        "standards_version": str(p.cell(7,2).value or "2022"),
    }
    zones = []
    for i, row in enumerate(wb["Zones"].iter_rows(min_row=2, values_only=True), 2):
        if not row[0] or str(row[0]).startswith("#"): continue
        zid = str(row[0]).strip()
        zones.append({"_row":i,"id":zid,"name":str(row[1] or zid),
                       "area":flt(row[2]),"height":flt(row[3],9.0),
                       "cond_type":str(row[4] or "Conditioned"),
                       "occ_type":str(row[5] or "") if len(row)>5 else "",
                       "floor":int(flt(row[6],1)) if len(row)>6 and row[6] not in (None,"") else 1})
    walls = []
    for i, row in enumerate(wb["Walls"].iter_rows(min_row=2, values_only=True), 2):
        if not row[0] or str(row[0]).startswith("#"): continue
        wid = str(row[0]).strip()
        walls.append({"_row":i,"id":wid,"zone_id":str(row[1] or "").strip(),
                       "name":str(row[2] or wid),"type":str(row[3] or "Exterior Wall"),
                       "orientation":str(row[4] or ""),"azimuth":azimuth(row[4]),
                       "area":flt(row[5]),"construction":str(row[6] or ""),
                       "adj_zone":str(row[7] or "").strip() if len(row)>7 else ""})
    openings = []
    for i, row in enumerate(wb["Openings"].iter_rows(min_row=2, values_only=True), 2):
        if not row[0] or str(row[0]).startswith("#"): continue
        oid = str(row[0]).strip()
        openings.append({"_row":i,"id":oid,"wall_id":str(row[1] or "").strip(),
                          "name":str(row[2] or oid),"type":str(row[3] or "Window"),
                          "area":flt(row[4]),
                          "ufactor":flt(row[5],None) if row[5] not in (None,"") else None,
                          "shgc":flt(row[6],None) if len(row)>6 and row[6] not in (None,"") else None})
    return {"project":project,"zones":zones,"walls":walls,"openings":openings}

def coerce(v):
    if v in (None, ""): return None
    try: return int(v) if "." not in str(v) else float(v)
    except: return v

@app.route("/")
def index(): return send_from_directory(SCRIPT_DIR, "viewer.html")

@app.route("/api/files")
def api_files(): return jsonify(find_xlsx())

@app.route("/api/data")
def api_data():
    fn = request.args.get("file") or (find_xlsx() or [""])[0]
    path = os.path.join(SCRIPT_DIR, fn)
    if not os.path.exists(path): return jsonify({"error": f"Not found: {fn}"}), 404
    try:
        d = read_data(path)
        d["_file"] = fn
        d["_mtime"] = os.path.getmtime(path)
        return jsonify(d)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/mtime")
def api_mtime():
    fn = request.args.get("file","")
    p = os.path.join(SCRIPT_DIR, fn)
    return jsonify({"mtime": os.path.getmtime(p) if os.path.exists(p) else 0})

@app.route("/api/update", methods=["POST"])
def api_update():
    d = request.json
    path = os.path.join(SCRIPT_DIR, d["file"])
    try:
        wb = openpyxl.load_workbook(path)
        wb[d["sheet"]].cell(int(d["row"]), int(d["col"]), coerce(d["value"]))
        wb.save(path)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/add_row", methods=["POST"])
def api_add_row():
    d = request.json
    path = os.path.join(SCRIPT_DIR, d["file"])
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb[d["sheet"]]
        # Find last non-empty, non-comment data row
        last = 1
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if v and not str(v).startswith("#"): last = r
        new_row = last + 1
        ws.insert_rows(new_row)
        for j, val in enumerate(d.get("values", []), 1):
            if val not in (None, ""): ws.cell(new_row, j, coerce(val))
        wb.save(path)
        return jsonify({"ok": True, "row": new_row})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/delete_row", methods=["POST"])
def api_delete_row():
    d = request.json
    path = os.path.join(SCRIPT_DIR, d["file"])
    try:
        wb = openpyxl.load_workbook(path)
        wb[d["sheet"]].delete_rows(int(d["row"]))
        wb.save(path)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/geometry")
def api_geometry():
    fn = request.args.get("file", "")
    base = os.path.splitext(os.path.join(SCRIPT_DIR, fn))[0]
    geo_path = base + "_geometry.json"
    if not os.path.exists(geo_path):
        return jsonify(None)
    try:
        import json
        with open(geo_path) as f:
            return jsonify(json.load(f))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/generate", methods=["POST"])
def api_generate():
    fn = request.json.get("file")
    xlsx = os.path.join(SCRIPT_DIR, fn)
    out = os.path.splitext(xlsx)[0] + ".xml"
    try:
        r = subprocess.run(
            [sys.executable, os.path.join(SCRIPT_DIR, "generate_gbxml.py"), xlsx, out],
            capture_output=True, text=True, timeout=30)
        return jsonify({"ok": r.returncode==0, "output": r.stdout+r.stderr,
                        "outfile": os.path.basename(out)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("="*50)
    print("  T24 3D Viewer  ->  http://localhost:5000")
    print("  Ctrl+C to stop")
    print("="*50)
    threading.Thread(
        target=lambda: (time.sleep(1.2), webbrowser.open("http://localhost:5000")),
        daemon=True).start()
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)
