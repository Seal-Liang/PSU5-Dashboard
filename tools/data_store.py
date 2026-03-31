import json
import os
import sys

import shutil
import threading

_file_lock = threading.Lock()

# Secure Data Storage to avoid overwrite by PyInstaller
appdata = os.environ.get('APPDATA')
if appdata:
    BASE_DIR = os.path.join(appdata, 'WeeklyDashboard')
else:
    BASE_DIR = os.path.join(os.path.expanduser('~'), '.weekly_dashboard')

os.makedirs(BASE_DIR, exist_ok=True)
DATA_FILE = os.path.join(BASE_DIR, 'data.json')

# Migration logic for existing users
old_data_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data.json')
if getattr(sys, 'frozen', False):
    old_data_file = os.path.join(os.path.dirname(sys.executable), 'data.json')

if not os.path.exists(DATA_FILE) and os.path.exists(old_data_file):
    try:
        shutil.copy2(old_data_file, DATA_FILE)
    except:
        pass

def load_data():
    with _file_lock:
        if not os.path.exists(DATA_FILE):
            return {
                "Settings": {
                    "departments": [],
                    "holidays": []
                },
                "BUMapping": [],
                "WeeklyData": []
            }
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except UnicodeDecodeError:
            try:
                with open(DATA_FILE, 'r', encoding='big5') as f:
                    # If this works, immediately resave it identically as UTF-8
                    data = json.load(f)
                    with open(DATA_FILE, 'w', encoding='utf-8') as out_f:
                        json.dump(data, out_f, ensure_ascii=False, indent=2)
                    return data
            except Exception:
                pass
                
        except Exception:
            pass
            
        # If we reached here, parsing failed completely. Backup and return fresh.
        try:
            import time
            os.rename(DATA_FILE, DATA_FILE + f".bak_{int(time.time())}")
        except:
            pass
        return { "Settings": { "departments": [], "holidays": [], "headcount_logs": [] }, "BUMapping": [], "WeeklyData": [] }

def save_data(data):
    with _file_lock:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

def save_weekly_data(weekly_obj):
    data = load_data()
    
    # Check if week_id exists
    week_id = weekly_obj.get("week_id")
    found_idx = -1
    for i, w in enumerate(data["WeeklyData"]):
        if w.get("week_id") == week_id:
            found_idx = i
            break
            
    if found_idx != -1:
        # Overwrite
        # Preserve working_days if the new object doesn't have it explicitly set by user, 
        # but the weekly_obj from parser has default 5.
        # Let's preserve existing working_days if present.
        existing_days = data["WeeklyData"][found_idx].get("working_days", 5)
        weekly_obj["working_days"] = existing_days
        data["WeeklyData"][found_idx] = weekly_obj
    else:
        # Append
        data["WeeklyData"].append(weekly_obj)
        
    save_data(data)

def save_bu_mapping(mapping_dict):
    data = load_data()
    # Convert dict to array of objects
    bu_array = []
    for code, details in mapping_dict.items():
        bu_array.append({
            "product_code": code,
            "bu": details["bu"],
            "bg": details["bg"],
            "product_name": details["product_name"]
        })
    data["BUMapping"] = bu_array
    save_data(data)

def update_settings(settings_update):
    data = load_data()
    if "departments" in settings_update:
        data["Settings"]["departments"] = settings_update["departments"]
    if "headcount_logs" in settings_update:
        data["Settings"]["headcount_logs"] = settings_update["headcount_logs"]
    if "holidays" in settings_update:
        data["Settings"]["holidays"] = settings_update["holidays"]
    save_data(data)
