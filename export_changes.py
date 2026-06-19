"""
export_changes.py
=================
Fetches pending work order changes from Railway and writes them to
WO_Changes.xlsx for Yardi writeback.

Run with: py export_changes.py
"""

import os
import requests
import openpyxl
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

RAILWAY_URL  = os.getenv("RAILWAY_URL", "").rstrip("/")
SYNC_API_KEY = os.getenv("SYNC_API_KEY", "harvest-workorder-2026-secure-key")
OUTPUT_PATH  = r"I:\PycharmProjects\WorkOrderBot\WO_Changes.xlsx"

PENDING_SHEET   = "Pending"
COMPLETED_SHEET = "Completed"

HEADERS = [
    "Change ID", "WO#", "Property", "Unit", "Brief Desc",
    "Field Changed", "Old Value", "New Value",
    "Changed By", "Changed At", "Assigned Tech"
]


def fetch_pending_changes():
    url = f"{RAILWAY_URL}/api/changes/export"
    headers = {"X-API-Key": SYNC_API_KEY}
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()


def mark_completed(ids: list):
    if not ids:
        return
    url = f"{RAILWAY_URL}/api/changes/mark_complete"
    headers = {
        "Content-Type": "application/json",
        "X-API-Key": SYNC_API_KEY,
    }
    resp = requests.post(url, headers=headers, json={"ids": ids}, timeout=30)
    resp.raise_for_status()
    return resp.json()


def load_or_create_workbook(path: str) -> openpyxl.Workbook:
    if os.path.exists(path):
        return openpyxl.load_workbook(path)
    wb = openpyxl.Workbook()
    # Create Pending sheet
    ws_pending = wb.active
    ws_pending.title = PENDING_SHEET
    ws_pending.append(HEADERS)
    # Create Completed sheet
    ws_completed = wb.create_sheet(COMPLETED_SHEET)
    ws_completed.append(HEADERS)
    return wb


def get_or_create_sheet(wb: openpyxl.Workbook, name: str) -> openpyxl.worksheet.worksheet.Worksheet:
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    ws.append(HEADERS)
    return ws


def change_to_row(change: dict) -> list:
    return [
        change.get("id"),
        change.get("wo_id"),
        change.get("property_name", ""),
        change.get("unit_number", ""),
        change.get("brief_desc", ""),
        change.get("field"),
        change.get("old_value", ""),
        change.get("new_value", ""),
        change.get("changed_by", ""),
        change.get("changed_at", ""),
        change.get("tech_name", ""),
    ]


def main():
    if not RAILWAY_URL:
        print("ERROR: RAILWAY_URL not set in .env")
        return

    print(f"Fetching pending changes from {RAILWAY_URL}...")
    try:
        changes = fetch_pending_changes()
    except requests.exceptions.RequestException as e:
        print(f"ERROR: Could not fetch changes — {e}")
        return

    if not changes:
        print("No pending changes found.")
        return

    print(f"Found {len(changes)} pending change(s).")

    wb = load_or_create_workbook(OUTPUT_PATH)
    ws_pending   = get_or_create_sheet(wb, PENDING_SHEET)
    ws_completed = get_or_create_sheet(wb, COMPLETED_SHEET)

    # Move existing Pending rows to Completed sheet first
    pending_rows = list(ws_pending.iter_rows(min_row=2, values_only=True))
    if pending_rows:
        print(f"Moving {len(pending_rows)} existing pending row(s) to Completed...")
        for row in pending_rows:
            ws_completed.append(list(row))
        # Clear pending sheet (keep header)
        for row in ws_pending.iter_rows(min_row=2):
            for cell in row:
                cell.value = None
        # Delete the blank rows
        ws_pending.delete_rows(2, ws_pending.max_row)

    # Write new pending changes to Pending sheet
    ids_to_mark = []
    for change in changes:
        ws_pending.append(change_to_row(change))
        ids_to_mark.append(change["id"])

    wb.save(OUTPUT_PATH)
    print(f"Saved {len(changes)} change(s) to: {OUTPUT_PATH} [{PENDING_SHEET}]")

    # Mark as completed in Railway
    print("Marking changes as completed in Railway...")
    try:
        result = mark_completed(ids_to_mark)
        print(f"Marked {result.get('marked', '?')} change(s) as completed.")
    except requests.exceptions.RequestException as e:
        print(f"WARNING: Could not mark changes as completed — {e}")
        print("Changes were written to Excel but remain 'pending' in Railway.")

    print("Done.")


if __name__ == "__main__":
    main()
