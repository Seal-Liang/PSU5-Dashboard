import json
import sys

try:
    with open('d:/CODE/PSU5/data.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
except Exception as e:
    print(f"Error reading JSON: {e}")
    sys.exit(1)

print("--- BU Mapping Keys ---")
mappings = data.get("BUMapping", [])
for m in mappings[:10]:
    print(repr(m.get("product_code")))

print("\n--- Weekly Data Product Codes ---")
wk_data = data.get("WeeklyData", [])
unmapped_count = 0
total_records = 0
for w in wk_data:
    for r in w.get("records", []):
        if r.get('bu') == 'Unknown':
            unmapped_count += 1
        total_records += 1

print(f"\nTotal Records: {total_records}")
print(f"Unmapped: {unmapped_count}")
    
print(f"\nUnmapped: {unmapped_count}, Mapped: {mapped_count}")

# Check for a specific match
if mappings and wk_data:
    m = mappings[0].get("product_code", "")
    w = wk_data[0].get("records", [])[0].get("product_code", "")
    print(f"\nComparing '{m}' and '{w}'")
    print(f"Match exactly? {m == w}")
    print(f"Match uppercase? {m.upper() == w.upper()}")

