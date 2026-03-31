import os
import sys
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from tools.excel_parser import parse_bu_mapping, parse_weekly_report
from tools.data_store import save_bu_mapping, save_weekly_data, load_data, save_data, update_settings

if getattr(sys, 'frozen', False):
    # PyInstaller bundle
    static_folder = os.path.join(sys._MEIPASS, 'static')
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Normal Python
    static_folder = 'static'
    BASE_DIR = os.path.dirname(__file__)

app = Flask(__name__, static_folder=static_folder, static_url_path='')

TMP_DIR = os.path.join(BASE_DIR, '.tmp')
os.makedirs(TMP_DIR, exist_ok=True)

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/api/upload_bu', methods=['POST'])
def upload_bu():
    bu_file = request.files.get('bu_mapping')
    if not bu_file or bu_file.filename == '':
        return jsonify({"error": "Missing or empty BU file."}), 400
        
    bu_path = os.path.join(TMP_DIR, secure_filename(bu_file.filename))
    bu_file.save(bu_path)
    
    try:
        mapping_dict = parse_bu_mapping(bu_path)
        # Convert all keys in mapping_dict to upper to guarantee robustness
        upper_mapping = {k.upper(): v for k, v in mapping_dict.items()}
        save_bu_mapping(upper_mapping)
        
        # Retroactively apply to existing weekly data
        data = load_data()
        for w in data.get("WeeklyData", []):
            for r in w.get("records", []):
                pc = str(r.get("product_code", "")).strip().upper()
                if pc in upper_mapping:
                    r["bu"] = upper_mapping[pc]["bu"]
                    r["bg"] = upper_mapping[pc]["bg"]
        save_data(data)

        return jsonify({"status": "success", "message": "BU Mapping processed and retroactively applied."})
    except Exception as e:
        return jsonify({"error": f"Error parsing BU file: {str(e)}"}), 500
    finally:
        try:
            if os.path.exists(bu_path): os.remove(bu_path)
        except: pass

@app.route('/api/upload_reports', methods=['POST'])
def upload_reports():
    reports = request.files.getlist('weekly_reports')
    if not reports:
        return jsonify({"error": "No files provided."}), 400
        
    data = load_data()
    bu_mapping_arr = data.get("BUMapping", [])
    mapping_dict = {str(m["product_code"]).strip().upper(): {"bu": m["bu"], "bg": m["bg"], "product_name": m["product_name"]} for m in bu_mapping_arr}
    
    results = []
    has_error = False
    
    for report in reports:
        if report.filename == '': continue
        path = os.path.join(TMP_DIR, secure_filename(report.filename))
        report.save(path)
        try:
            weekly_data = parse_weekly_report(path, mapping_dict)
            save_weekly_data(weekly_data)
            results.append({"file": report.filename, "status": "success"})
        except Exception as e:
            has_error = True
            results.append({"file": report.filename, "error": str(e)})
        finally:
            try:
                if os.path.exists(path): os.remove(path)
            except: pass
            
    if has_error:
        return jsonify({"status": "partial_error", "results": results}), 207
    return jsonify({"status": "success", "results": results})

@app.route('/api/upload_holidays', methods=['POST'])
def upload_holidays():
    file = request.files.get('holiday_file')
    if not file or file.filename == '':
        return jsonify({"error": "No file provided."}), 400
        
    path = os.path.join(TMP_DIR, secure_filename(file.filename))
    file.save(path)
    try:
        import pandas as pd
        if path.endswith('.csv'):
            df = pd.read_csv(path)
        else:
            df = pd.read_excel(path)
            
        new_holidays = []
        for _, row in df.iterrows():
            d = row.get("Date")
            n = row.get("Name", "Holiday")
            if pd.isna(d): continue
            
            try:
                date_str = pd.to_datetime(d).strftime('%Y-%m-%d')
                new_holidays.append({"date": date_str, "name": str(n)})
            except:
                pass
                
        data = load_data()
        existing = data["Settings"].get("holidays", [])
        
        # Deduplicate based on date
        existing_dates = set(h["date"] if isinstance(h, dict) else h for h in existing)
        added = 0
        for h in new_holidays:
            if h["date"] not in existing_dates:
                existing.append(h)
                existing_dates.add(h["date"])
                added += 1
                
        existing.sort(key=lambda x: x["date"] if isinstance(x, dict) else x)
        data["Settings"]["holidays"] = existing
        save_data(data)
        
        return jsonify({"status": "success", "message": f"Added {added} new holidays."})
    except Exception as e:
        return jsonify({"error": f"Error parsing holiday file: {str(e)}"}), 500
    finally:
        try:
            if os.path.exists(path): os.remove(path)
        except: pass

@app.route('/api/upload_logs', methods=['POST'])
def upload_logs():
    file = request.files.get('log_file')
    target_team = request.form.get('target_team')
    
    if not file or file.filename == '':
        return jsonify({"error": "No file provided."}), 400
        
    path = os.path.join(TMP_DIR, secure_filename(file.filename))
    file.save(path)
    try:
        import pandas as pd
        if path.lower().endswith('.csv'):
            try:
                df = pd.read_csv(path)
            except:
                df = pd.read_csv(path, encoding='big5', errors='replace')
        else:
            try:
                df = pd.read_excel(path)
            except:
                df = pd.read_csv(path, encoding='big5', errors='replace')

        data = load_data()
        logs_data = data["Settings"].get("headcount_logs", [])
        team_dict = {l["department"]: l for l in logs_data}
        
        added = 0
        for _, row in df.iterrows():
            d = row.get("Date")
            t_col = row.get("Team")
            raw_type = str(row.get("Type", "")).strip()
            name = str(row.get("Name", "From file")).strip()
            
            if pd.isna(d): continue
            team_to_use = str(t_col).strip() if pd.notna(t_col) and str(t_col).strip() else target_team
            if not team_to_use: continue
                
            try:
                date_str = pd.to_datetime(d).strftime('%Y-%m-%d')
            except:
                continue
                
            delta = 1
            if 'Leave' in raw_type or 'Depart' in raw_type or 'NoTracking' in raw_type or '離職' in raw_type or '留停' in raw_type:
                delta = -1
                
            entry = {"date": date_str, "type": raw_type, "delta": delta, "name": name}
            
            if team_to_use not in team_dict:
                team_dict[team_to_use] = {"department": team_to_use, "initial": 0, "logs": []}
                
            existing = team_dict[team_to_use]["logs"]
            if not any(e["date"] == date_str and e["name"] == name and e["type"] == raw_type for e in existing):
                existing.append(entry)
                added += 1
                
        for tObj in team_dict.values():
            tObj["logs"].sort(key=lambda x: x["date"])
            
        data["Settings"]["headcount_logs"] = list(team_dict.values())
        save_data(data)
        
        return jsonify({"status": "success", "message": f"Added {added} log entries."})
    except Exception as e:
        return jsonify({"error": f"Error parsing log file: {str(e)}"}), 500
    finally:
        try:
            if os.path.exists(path): os.remove(path)
        except: pass

@app.route('/api/data', methods=['GET'])
def get_data():
    return jsonify(load_data())

@app.route('/api/settings', methods=['POST'])
def set_settings():
    data = request.json
    if not data:
        return jsonify({"error": "No JSON payload provided."}), 400
    update_settings(data)
    return jsonify({"status": "success"})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
