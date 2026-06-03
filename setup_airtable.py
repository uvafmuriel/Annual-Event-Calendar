import os, urllib.request, urllib.error, json

TOKEN   = os.environ["AIRTABLE_TOKEN"]
BASE_ID = os.environ["AIRTABLE_BASE_ID"]
URL     = f"https://api.airtable.com/v0/meta/bases/{BASE_ID}/tables"
HEADERS = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}

# First test the token works at all
print("Testing token...")
try:
    req = urllib.request.Request(
        "https://api.airtable.com/v0/meta/whoami",
        headers={"Authorization": f"Bearer {TOKEN}"}
    )
    with urllib.request.urlopen(req) as r:
        data = json.loads(r.read())
        print(f"  Token valid — user ID: {data.get('id', 'unknown')}")
        print(f"  Scopes: {data.get('scopes', [])}")
except urllib.error.HTTPError as e:
    print(f"  Token test FAILED: {e.code} — {e.read().decode()}")

# Test base access
print(f"\nTesting base access ({BASE_ID})...")
try:
    req = urllib.request.Request(
        f"https://api.airtable.com/v0/meta/bases/{BASE_ID}/tables",
        headers={"Authorization": f"Bearer {TOKEN}"}
    )
    with urllib.request.urlopen(req) as r:
        data = json.loads(r.read())
        print(f"  Base accessible — {len(data.get('tables',[]))} existing tables")
except urllib.error.HTTPError as e:
    print(f"  Base access FAILED: {e.code} — {e.read().decode()}")
