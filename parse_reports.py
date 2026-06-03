"""
CGZ Board — Weekly XLSX Parser → Airtable
Reads XLSX files from weekly-reports/, parses events for the upcoming Mon-Fri,
clears old Weekly Events records for that week, and writes fresh ones to Airtable.
Also updates the Week Of field on the active Announcements row.
"""

import json
import os
import glob
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime, timedelta, date
import openpyxl

# ── CONFIG ────────────────────────────────────────────────────────────────────
TOKEN       = os.environ.get("AIRTABLE_TOKEN", "")
BASE_ID     = os.environ.get("AIRTABLE_BASE_ID", "")
REPORTS_DIR = "weekly-reports"

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type":  "application/json"
}

SKIP_TYPES       = {"Maintenance"}
SKIP_NAME_STARTS = ("maintenance hold", "sa event mgt hold")
DAY_NAMES        = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

# ── AIRTABLE HELPERS ──────────────────────────────────────────────────────────
def at_request(method, path, data=None, params=""):
    url = f"https://api.airtable.com/v0/{BASE_ID}/{urllib.parse.quote(path)}{'?' + params if params else ''}"
    body = json.dumps(data).encode() if data else None
    req = urllib.request.Request(url, data=body, headers=HEADERS, method=method)
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())

def at_get(path, params=""): return at_request("GET", path, params=params)
def at_post(path, data):     return at_request("POST", path, data)
def at_patch(path, data):    return at_request("PATCH", path, data)

def at_delete_batch(path, ids):
    for i in range(0, len(ids), 10):
        batch = ids[i:i+10]
        params = "&".join(f"records[]={r}" for r in batch)
        url = f"https://api.airtable.com/v0/{BASE_ID}/{urllib.parse.quote(path)}?{params}"
        req = urllib.request.Request(url, headers=HEADERS, method="DELETE")
        with urllib.request.urlopen(req) as r:
            pass

def get_all_records(table):
    records, offset = [], None
    while True:
        params = f"offset={offset}" if offset else ""
        data = at_get(table, params)
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset: break
    return records

# ── DATE HELPERS ──────────────────────────────────────────────────────────────
def get_next_week_range():
    today = date.today()
    wd = today.weekday()
    days_to_monday = (7 - wd) if wd >= 4 else -wd
    monday = today + timedelta(days=days_to_monday)
    return monday, monday + timedelta(days=4)

def week_of_label(monday):
    friday = monday + timedelta(days=4)
    if monday.month == friday.month:
        return f"{monday.strftime('%B %-d')} \u2013 {friday.strftime('%-d, %Y')}"
    return f"{monday.strftime('%B %-d')} \u2013 {friday.strftime('%B %-d, %Y')}"

def fmt_time(dt):
    if not dt or not hasattr(dt, 'strftime'): return ""
    return dt.strftime("%-I:%M %p")

def clean_name(name):
    name = str(name or "").strip()
    return name[:42].rstrip() + "\u2026" if len(name) > 45 else name

# ── XLSX PARSER ───────────────────────────────────────────────────────────────
def parse_xlsx(filepath):
    events = []
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        if "Reservations" not in wb.sheetnames:
            print(f"  No Reservations sheet in {filepath}")
            return events
        ws = wb["Reservations"]
        rows = list(ws.iter_rows(values_only=True))
        header_row, data_start = None, 0
        for i, row in enumerate(rows):
            if row and row[0] == "Event Name":
                header_row, data_start = row, i + 1
                break
        if not header_row: return events
        col = {str(h).strip(): i for i, h in enumerate(header_row) if h}
        for row in rows[data_start:]:
            if not row or not row[col.get("Event Name", 0)]: continue
            name       = row[col.get("Event Name", 0)]
            day        = row[col.get("Day", 1)]
            start      = row[col.get("Event Start", 4)]
            end        = row[col.get("Event End", 5)]
            location   = row[col.get("Location", 8)]
            event_type = row[col.get("Event Type", 17)]
            state      = row[col.get("Event State", 18)]
            if event_type in SKIP_TYPES: continue
            if str(name).strip().lower().startswith(SKIP_NAME_STARTS): continue
            if state not in ("Confirmed", "Tentative"): continue
            if not day or not hasattr(day, 'date'): continue
            events.append({
                "date":     day.date(),
                "name":     clean_name(name),
                "time":     fmt_time(start),
                "end_time": fmt_time(end),
                "location": str(location or "").strip(),
                "type":     str(event_type or "").strip(),
                "state":    str(state or "").strip(),
                "source":   os.path.basename(filepath),
            })
        wb.close()
    except Exception as e:
        print(f"  Error parsing {filepath}: {e}")
    return events

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print("CGZ Board Parser -> Airtable")
    xlsx_files = glob.glob(os.path.join(REPORTS_DIR, "*.xlsx"))
    if not xlsx_files:
        print(f"No XLSX files in {REPORTS_DIR}/ -- nothing to do")
        return

    print(f"Found {len(xlsx_files)} file(s)")
    all_events = []
    for f in xlsx_files:
        print(f"Parsing {os.path.basename(f)}...")
        evs = parse_xlsx(f)
        print(f"  -> {len(evs)} events")
        all_events.extend(evs)

    monday, friday = get_next_week_range()
    week_label = week_of_label(monday)
    print(f"\nTarget week: {monday} -> {friday} ({week_label})")

    week_events = sorted(
        [e for e in all_events if monday <= e["date"] <= friday],
        key=lambda e: (e["date"], e["time"])
    )
    print(f"Events this week: {len(week_events)}")

    print("\nClearing old Airtable records for this week...")
    existing = get_all_records("Weekly Events")
    to_delete = [r["id"] for r in existing if r["fields"].get("Week Of") == week_label]
    if to_delete:
        at_delete_batch("Weekly Events", to_delete)
        print(f"  Deleted {len(to_delete)} old records")

    records = [{"fields": {
        "Event Name":  ev["name"],
        "Date":        ev["date"].isoformat(),
        "Day":         DAY_NAMES[ev["date"].weekday()],
        "Start Time":  ev["time"],
        "End Time":    ev["end_time"],
        "Location":    ev["location"],
        "Event Type":  ev["type"],
        "Status":      ev["state"],
        "Source File": ev["source"],
        "Week Of":     week_label,
    }} for ev in week_events]

    created = 0
    for i in range(0, len(records), 10):
        result = at_post("Weekly Events", {"records": records[i:i+10]})
        created += len(result.get("records", []))
    print(f"  Created {created} event records")

    ann = get_all_records("Announcements")
    active = [r for r in ann if r["fields"].get("Active")]
    if active:
        at_patch(f"Announcements/{active[0]['id']}", {"fields": {"Week Of": week_label}})
        print(f"  Updated Announcements week to: {week_label}")
    else:
        at_post("Announcements", {"records": [{"fields": {"Week Of": week_label, "Active": True}}]})
        print(f"  Created Announcements row for: {week_label}")

    print("\nDone!")

if __name__ == "__main__":
    main()
