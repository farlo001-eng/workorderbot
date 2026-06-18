import os
import json
import uuid
from datetime import datetime
from flask import Flask, jsonify, request, send_from_directory
from dotenv import load_dotenv
from database import init_db, get_db

load_dotenv()
app = Flask(__name__, static_folder="static", template_folder="templates")

SYNC_API_KEY = os.getenv("SYNC_API_KEY", "harvest-workorder-2026-secure-key")

# Initialize DB on startup
init_db()


# --- Page routes ---

@app.route("/")
def index():
    return send_from_directory("templates", "dashboard.html")


@app.route("/workorder")
def workorder():
    return send_from_directory("templates", "workorder.html")


@app.route("/dashboard")
def dashboard():
    return send_from_directory("templates", "dashboard.html")


# --- API: List work orders ---

@app.route("/api/workorders", methods=["GET"])
def get_workorders():
    property_code = request.args.get("property")
    status = request.args.get("status")
    employee = request.args.get("employee")

    conn = get_db()
    c = conn.cursor()

    query = "SELECT * FROM work_orders WHERE 1=1"
    params = []

    if property_code:
        query += " AND property_code = ?"
        params.append(property_code)
    if status:
        query += " AND status = ?"
        params.append(status)
    if employee:
        query += " AND employee = ?"
        params.append(employee)

    # Default sort: ai_priority ASC (1=highest), then days_open DESC
    query += " ORDER BY COALESCE(ai_priority, 99) ASC, COALESCE(days_open, 0) DESC"

    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    return jsonify([dict(r) for r in rows])


# --- API: Create work order manually ---

@app.route("/api/workorder/create", methods=["POST"])
def create_workorder():
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400

    wo_id = data.get("id") or str(uuid.uuid4())
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT OR IGNORE INTO work_orders
        (id, property_code, property_name, unit_number, description, brief_desc,
         category, priority, status, created_date, tech_notes, photos, source,
         days_open, ai_priority, ai_reason)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'manual', 0, NULL, NULL)
    """, (
        wo_id,
        data.get("property_code", ""),
        data.get("property_name", ""),
        data.get("unit_number", ""),
        data.get("description", ""),
        data.get("brief_desc", ""),
        data.get("category", ""),
        data.get("priority", "Routine"),
        data.get("status", "open"),
        now,
        data.get("tech_notes", ""),
        json.dumps([]),
    ))
    conn.commit()
    conn.close()

    return jsonify({"success": True, "id": wo_id})


# --- API: Update work order (tech updates status, notes, photos) ---

@app.route("/api/workorder/update", methods=["POST"])
def update_workorder():
    data = request.json
    if not data or "id" not in data:
        return jsonify({"error": "Missing work order id"}), 400

    wo_id = data["id"]
    conn = get_db()
    c = conn.cursor()

    fields = []
    params = []

    if "status" in data:
        fields.append("status = ?")
        params.append(data["status"])
        if data["status"] == "complete":
            fields.append("completed_date = ?")
            params.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    if "tech_notes" in data:
        fields.append("tech_notes = ?")
        params.append(data["tech_notes"])

    if "photos" in data:
        fields.append("photos = ?")
        params.append(json.dumps(data["photos"]))

    if not fields:
        return jsonify({"error": "No fields to update"}), 400

    params.append(wo_id)
    c.execute(f"UPDATE work_orders SET {', '.join(fields)} WHERE id = ?", params)
    conn.commit()
    conn.close()

    return jsonify({"success": True})


# --- API: Sync from Yardi (requires API key) ---

@app.route("/workorders/sync", methods=["POST"])
def sync_workorders():
    api_key = request.headers.get("X-API-Key")
    if api_key != SYNC_API_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json
    if not data or "work_orders" not in data:
        return jsonify({"error": "No work_orders payload"}), 400

    work_orders = data["work_orders"]
    inserted = 0
    updated = 0

    conn = get_db()
    c = conn.cursor()

    for wo in work_orders:
        wo_id = str(wo.get("id", "")).strip()
        if not wo_id:
            continue

        # Check if exists and preserve tech_notes + photos
        c.execute("SELECT tech_notes, photos FROM work_orders WHERE id = ?", (wo_id,))
        existing = c.fetchone()

        if existing:
            # UPDATE — preserve tech_notes and photos from field tech
            c.execute("""
                UPDATE work_orders SET
                    property_code = ?,
                    property_name = ?,
                    unit_number = ?,
                    description = ?,
                    brief_desc = ?,
                    category = ?,
                    priority = ?,
                    status = ?,
                    created_date = ?,
                    scheduled_date = ?,
                    completed_date = ?,
                    employee = ?,
                    actual_start = ?,
                    actual_finish = ?,
                    actual_hours = ?,
                    days_open = ?,
                    ai_priority = ?,
                    ai_reason = ?,
                    source = COALESCE(source, 'yardi')
                WHERE id = ?
            """, (
                wo.get("property_code"), wo.get("property_name"),
                wo.get("unit_number"), wo.get("description"),
                wo.get("brief_desc"), wo.get("category"),
                wo.get("priority"), wo.get("status"),
                wo.get("created_date"), wo.get("scheduled_date"),
                wo.get("completed_date"), wo.get("employee"),
                wo.get("actual_start"), wo.get("actual_finish"),
                wo.get("actual_hours"), wo.get("days_open"),
                wo.get("ai_priority"), wo.get("ai_reason"),
                wo_id
            ))
            updated += 1
        else:
            # INSERT new
            c.execute("""
                INSERT INTO work_orders
                (id, property_code, property_name, unit_number, description, brief_desc,
                 category, priority, status, created_date, scheduled_date, completed_date,
                 employee, actual_start, actual_finish, actual_hours, tech_notes, photos,
                 source, days_open, ai_priority, ai_reason)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', '[]', 'yardi', ?, ?, ?)
            """, (
                wo_id, wo.get("property_code"), wo.get("property_name"),
                wo.get("unit_number"), wo.get("description"), wo.get("brief_desc"),
                wo.get("category"), wo.get("priority"), wo.get("status"),
                wo.get("created_date"), wo.get("scheduled_date"), wo.get("completed_date"),
                wo.get("employee"), wo.get("actual_start"), wo.get("actual_finish"),
                wo.get("actual_hours"), wo.get("days_open"),
                wo.get("ai_priority"), wo.get("ai_reason")
            ))
            inserted += 1

    conn.commit()
    conn.close()

    return jsonify({"success": True, "inserted": inserted, "updated": updated})


# --- API: Properties reference list ---

@app.route("/api/properties", methods=["GET"])
def get_properties():
    properties = [
        {"code": "79800", "name": "CityCenter Place",        "tech": "Antonio Smith",   "tech_key": "asmith"},
        {"code": "79200", "name": "Shake House",              "tech": "Antonio Smith",   "tech_key": "asmith"},
        {"code": "75000", "name": "Grace Gardens",            "tech": "Antonio Smith",   "tech_key": "asmith"},
        {"code": "78200", "name": "Nova Highland",            "tech": "Antonio Smith",   "tech_key": "asmith"},
        {"code": "78600", "name": "Stadium I",                "tech": "Antonio Smith",   "tech_key": "asmith"},
        {"code": "78700", "name": "Stadium II",               "tech": "Antonio Smith",   "tech_key": "asmith"},
        {"code": "79100", "name": "The Nolan",                "tech": "Antonio Smith",   "tech_key": "asmith"},
        {"code": "78800", "name": "Uptown Villas",            "tech": "Antonio Smith",   "tech_key": "asmith"},
        {"code": "73800", "name": "Windsor",                  "tech": "Antonio Smith",   "tech_key": "asmith"},
        {"code": "74100", "name": "Cahaba Brook",             "tech": "Daniel Hall",     "tech_key": "dhall"},
        {"code": "78400", "name": "Eight on 16th",            "tech": "Daniel Hall",     "tech_key": "dhall"},
        {"code": "77700", "name": "Greendale",                "tech": "Daniel Hall",     "tech_key": "dhall"},
        {"code": "76400", "name": "Redtop at Greensprings",   "tech": "Daniel Hall",     "tech_key": "dhall"},
        {"code": "73500", "name": "The William",              "tech": "Daniel Hall",     "tech_key": "dhall"},
        {"code": "76700", "name": "The Woodley",              "tech": "Daniel Hall",     "tech_key": "dhall"},
        {"code": "76500", "name": "V Apartments",             "tech": "Daniel Hall",     "tech_key": "dhall"},
        {"code": "73600", "name": "Ellys",                    "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "77900", "name": "Ascent on 34th",           "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "79400", "name": "Central Station",          "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "78950", "name": "Drake",                    "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "78900", "name": "Madrid",                   "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "79600", "name": "Places at Red Rocks",      "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "71300", "name": "Rhodes Park",              "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "79700", "name": "Silver Oaks",              "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "79500", "name": "Steel City Flats",         "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "77500", "name": "Sycamore Manor",           "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "79300", "name": "Ivory House",              "tech": "Jonathan Morton", "tech_key": "morton"},
        {"code": "73700", "name": "Murals on Niazuma",        "tech": "Jonathan Morton", "tech_key": "morton"},
    ]
    return jsonify(properties)


# --- Serve logo ---

@app.route("/harvest-logo.jpg")
def logo():
    return send_from_directory("static", "harvest-logo.jpg")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
