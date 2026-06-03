import os, urllib.request, json

TOKEN   = os.environ["AIRTABLE_TOKEN"]
BASE_ID = os.environ["AIRTABLE_BASE_ID"]
URL     = f"https://api.airtable.com/v0/meta/bases/{BASE_ID}/tables"
HEADERS = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}

def create_table(name, fields):
    print(f"Creating: {name}...")
    body = json.dumps({"name": name, "fields": fields}).encode()
    req = urllib.request.Request(URL, data=body, headers=HEADERS, method="POST")
    try:
        with urllib.request.urlopen(req) as r:
            result = json.loads(r.read())
            print(f"  Done: {result['id']}")
    except Exception as e:
        print(f"  Error: {e}")

create_table("Weekly Events", [
    {"name": "Event Name",  "type": "singleLineText"},
    {"name": "Date",        "type": "date", "options": {"dateFormat": {"name": "us"}}},
    {"name": "Day",         "type": "singleSelect", "options": {"choices": [
        {"name": "Monday"}, {"name": "Tuesday"}, {"name": "Wednesday"},
        {"name": "Thursday"}, {"name": "Friday"}, {"name": "Saturday"}, {"name": "Sunday"}
    ]}},
    {"name": "Start Time",  "type": "singleLineText"},
    {"name": "End Time",    "type": "singleLineText"},
    {"name": "Location",    "type": "singleLineText"},
    {"name": "Event Type",  "type": "singleLineText"},
    {"name": "Status",      "type": "singleLineText"},
    {"name": "Source File", "type": "singleLineText"},
    {"name": "Week Of",     "type": "singleLineText"},
])

create_table("Announcements", [
    {"name": "Week Of",             "type": "singleLineText"},
    {"name": "Lunch Available",     "type": "checkbox",      "options": {"icon": "check", "color": "greenBright"}},
    {"name": "Lunch What",          "type": "singleLineText"},
    {"name": "Lunch Location",      "type": "singleLineText"},
    {"name": "Lunch Until",         "type": "singleLineText"},
    {"name": "Work Orders",         "type": "multilineText"},
    {"name": "Construction Alerts", "type": "multilineText"},
    {"name": "Active",              "type": "checkbox", "options": {"icon": "check", "color": "blueBright"}},
])
print("Done!")
