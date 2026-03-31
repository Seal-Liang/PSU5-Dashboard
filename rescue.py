import json, os, glob

appdata = os.environ.get('APPDATA')
base_dir = os.path.join(appdata, 'WeeklyDashboard')
bak_files = glob.glob(os.path.join(base_dir, 'data.json.bak_*'))
if not bak_files:
    print("No backup files found.")
    exit(1)

bak_file = sorted(bak_files)[-1]
print(f"Rescuing {bak_file}")

with open(bak_file, 'rb') as f:
    content = f.read()

content_str = content.decode('utf-8', errors='ignore')

# Fix the specific syntax error
# "project_name": "設計中心部門會議_x0028_Design_x0020_"department": "MBDC-VCPD",
content_str = content_str.replace('"設計中心部門會議_x0028_Design_x0020_"department"', '"設計中心部門會議_x0028_Design_x0020_",\n          "department"')

try:
    data = json.loads(content_str, strict=False)
    print(f"Successfully loaded JSON! Found {len(data['WeeklyData'])} weeks.")
    
    # Write repaired data to data.json
    data_file = os.path.join(base_dir, 'data.json')
    with open(data_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("Rescue complete! Saved to data.json")
except Exception as e:
    print(f"Failed to load JSON after patch: {e}")
    import re
    m = re.search(r'char\s+(\d+)', str(e))
    if m:
        idx = int(m.group(1))
        print("Context:")
        print(repr(content_str[max(0, idx-40):min(len(content_str), idx+40)]))
