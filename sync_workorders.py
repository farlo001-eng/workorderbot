import os
import json
import requests
import openpyxl
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

RAILWAY_URL = os.getenv("RAILWAY_URL", "").rstrip("/")
SYNC_API_KEY = os.getenv("SYNC_API_KEY", "harvest-workorder-2026-secure-key")
EXCEL_PATH = os.getenv("EXCEL_PATH", "WorkOrders.xlsx")

# Exact column names from Yardi export (Row 1 headers)
# WO#, Property, Property Name, Unit, Priority, Status, Category,
# Brief Desc., Problem Description, Technician Notes, Call Date,
# Scheduled Date, Completed Date, Employee, Actual Start, Actual Finish,
# Actual Hours, Amount, Call Date Parsed, Days Open, AI Priority, AI Reason

def parse_date(val):
    """Convert Excel serial date or string to YYYY-MM-DD string."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        # Excel serial date — convert
        try:
            from datetime import date
            delta = date.fromordinal(int(val) + 693594 - 1)  # Excel epoch offset
            return delta.strftime("%Y-%m-%d")
        except Exception:
            return str(val)
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    return str(val).strip() if val else None


def main():
    if not RAILWAY_URL:
        print("ERROR: RAILWAY_URL not set in .env")
        return

    if not os.path.exists(EXCEL_PATH):
        print(f"ERROR: Excel file not found: {EXCEL_PATH}")
        return

    print(f"Reading {EXCEL_PATH}...")
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)

    # Use first sheet (should be "Open WO Data" or similar)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    if not rows:
        print("ERROR: Excel file is empty.")
        return

    headers = [str(h).strip() if h else "" for h in rows[0]]
    print(f"Columns found: {headers}")

    # Map column names to indices
    def col(name):
        try:
            return headers.index(name)
        except ValueError:
            return None

    work_orders = []
    skipped = 0

    for row in rows[1:]:
        wo_id = row[col("WO#")] if col("WO#") is not None else None
        if not wo_id:
            skipped += 1
            continue

        wo = {
            "id":             str(wo_id).strip(),
            "property_code":  str(row[col("Property")]).strip()      if col("Property") is not None and row[col("Property")] else "",
            "property_name":  str(row[col("Property Name")]).strip() if col("Property Name") is not None and row[col("Property Name")] else "",
            "unit_number":    str(row[col("Unit")]).strip()           if col("Unit") is not None and row[col("Unit")] else "",
            "description":    str(row[col("Problem Description")]).strip() if col("Problem Description") is not None and row[col("Problem Description")] else "",
            "brief_desc":     str(row[col("Brief Desc.")]).strip()    if col("Brief Desc.") is not None and row[col("Brief Desc.")] else "",
            "category":       str(row[col("Category")]).strip()       if col("Category") is not None and row[col("Category")] else "",
            "priority":       str(row[col("Priority")]).strip()       if col("Priority") is not None and row[col("Priority")] else "",
            "status":         str(row[col("Status")]).strip()         if col("Status") is not None and row[col("Status")] else "open",
            "created_date":   parse_date(row[col("Call Date")])       if col("Call Date") is not None else None,
            "scheduled_date": parse_date(row[col("Scheduled Date")])  if col("Scheduled Date") is not None else None,
            "completed_date": parse_date(row[col("Completed Date")])  if col("Completed Date") is not None else None,
            "employee":       str(row[col("Employee")]).strip()       if col("Employee") is not None and row[col("Employee")] else "",
            "actual_start":   parse_date(row[col("Actual Start")])    if col("Actual Start") is not None else None,
            "actual_finish":  parse_date(row[col("Actual Finish")])   if col("Actual Finish") is not None else None,
            "actual_hours":   row[col("Actual Hours")]                if col("Actual Hours") is not None else 0,
            "days_open":      int(row[col("Days Open")]) if col("Days Open") is not None and row[col("Days Open")] is not None else 0,
            "ai_priority":    int(row[col("AI Priority")]) if col("AI Priority") is not None and row[col("AI Priority")] is not None else None,
            "ai_reason":      str(row[col("AI Reason")]).strip()      if col("AI Reason") is not None and row[col("AI Reason")] else "",
        }
        work_orders.append(wo)

    wb.close()
    print(f"Parsed {len(work_orders)} work orders. Skipped {skipped} rows with no WO#.")

    if not work_orders:
        print("Nothing to sync.")
        return

    url = f"{RAILWAY_URL}/workorders/sync"
    headers_req = {
        "Content-Type": "application/json",
        "X-API-Key": SYNC_API_KEY,
    }
    payload = {"work_orders": work_orders}

    print(f"POSTing to {url}...")
    try:
        resp = requests.post(url, headers=headers_req, json=payload, timeout=60)
        resp.raise_for_status()
        result = resp.json()
        print(f"Sync complete. Inserted: {result.get('inserted', '?')}, Updated: {result.get('updated', '?')}")
    except requests.exceptions.RequestException as e:
        print(f"ERROR: Sync failed — {e}")


if __name__ == "__main__":
    main()
