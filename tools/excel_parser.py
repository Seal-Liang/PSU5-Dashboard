import pandas as pd
import math
import re

def parse_bu_mapping(filepath):
    """
    Parses the BU classification Excel/CSV file.
    Returns a dictionary mapping 'Product Code' to mapping details.
    """
    if filepath.lower().endswith('.csv'):
        try:
            df = pd.read_csv(filepath)
        except Exception:
            try:
                df = pd.read_csv(filepath, encoding='big5')
            except Exception:
                df = pd.read_csv(filepath, encoding='utf-8', errors='replace')
    else:
        try:
            xl = pd.ExcelFile(filepath)
            df = xl.parse(xl.sheet_names[0])
            xl.close()
        except ValueError:
            try:
                df = pd.read_csv(filepath)
            except Exception:
                df = pd.read_csv(filepath, encoding='big5', errors='replace')
    
    # Expected columns: ['Unnamed: 0', 'BG', 'BU', 'Product Code', 'Product Name', '聯絡窗口...']
    mapping = {}
    
    for _, row in df.iterrows():
        prod_code = str(row.get('Product Code', '')).strip()
        if pd.isna(row.get('Product Code')) or not prod_code or prod_code == 'nan':
            continue
            
        mapping[prod_code] = {
            'bg': str(row.get('BG', '')).strip(),
            'bu': str(row.get('BU', '')).strip(),
            'product_name': str(row.get('Product Name', '')).strip()
        }
        
    return mapping

def parse_weekly_report(filepath, bu_mapping, filename=""):
    """
    Parses the Weekly Data Excel file.
    """
    import os
    filename_upper = filename.upper() if filename else os.path.basename(filepath).upper()
    
    if "MBDC-ID" in filename_upper:
        source_type = "ID_DETAIL"
    elif "MBDC-UX" in filename_upper:
        source_type = "UX_DETAIL"
    else:
        source_type = "GENERAL"

    xl = pd.ExcelFile(filepath)
    sheets = xl.sheet_names
    
    # Extract week info from the first sheet (e.g., 'Total' or the first department)
    df_meta = xl.parse(sheets[0], header=None, nrows=1)
    meta_str = str(df_meta.iloc[0, 0])  # '2025WK31 : 2025/07/28'
    
    week_id = ""
    date_str = ""
    
    if " : " in meta_str:
        week_id, date_str = meta_str.split(" : ", 1)
        week_id = week_id.strip()
        date_str = date_str.strip()
    else:
        # Fallback or regex if format slightly differs
        m = re.match(r'(.*?WK\d+)\s*[:\-]\s*(.*)', meta_str)
        if m:
            week_id = m.group(1).strip()
            date_str = m.group(2).strip()
        else:
            week_id = meta_str

    records = []
    
    # Track hours per (proj_name, pic, raw_bu) to calculate difference against Total sheet
    sub_team_hours_map = {}
    
    # Iterate through sheets representing departments
    skip_sheets = {'Total', 'AMS'}
    
    for sheet in sheets:
        if sheet in skip_sheets:
            continue
            
        # Parse data assuming headers are at row 5 (0-indexed)
        try:
            df = xl.parse(sheet, header=5)
        except Exception:
            continue # If sheet doesn't have enough rows
            
        # Clean up column names by stripping spaces
        df.columns = [str(c).strip() for c in df.columns]
        
        required_cols = ['PROJECT NAME', 'PIC', 'BU', 'Hours']
        if not all(col in df.columns for col in required_cols):
            # Try to handle case where header was not at row 5 in this specific sheet
            continue
            
        for _, row in df.iterrows():
            proj_name = row.get('PROJECT NAME')
            if pd.isna(proj_name) or str(proj_name).strip() == '':
                continue
                
            hours = row.get('Hours')
            if pd.isna(hours) or hours == 0:
                continue
                
            raw_bu = str(row.get('BU')).strip()
            pic = str(row.get('PIC')).strip() if not pd.isna(row.get('PIC')) else "Unknown"

            # Track hours logged in sub-teams
            key = (str(proj_name).strip(), pic, raw_bu)
            sub_team_hours_map[key] = sub_team_hours_map.get(key, 0) + float(hours)

            # Support multiple product codes separated by commas (e.g. "BZHY, ABCD")
            # Hours are split evenly across all codes.
            prod_codes = [c.strip().upper() for c in raw_bu.split(',') if c.strip()]
            if not prod_codes:
                prod_codes = [raw_bu.upper()]
            split_hours = float(hours) / len(prod_codes)

            for prod_code in prod_codes:
                # Use mapping if available. The Excel 'BU' column is actually the Product Code
                mapped_bu = "Unknown"
                mapped_bg = "Unknown"
                
                if prod_code in bu_mapping:
                    mapped_bu = bu_mapping[prod_code]['bu']
                    mapped_bg = bu_mapping[prod_code]['bg']

                mapped_department = sheet
                # If sheet name does not include MBDC, it is an individual. Map to parent team.
                if "MBDC" not in sheet.upper():
                    if source_type == "ID_DETAIL":
                        mapped_department = "MBDC-ID"
                    elif source_type == "UX_DETAIL":
                        mapped_department = "MBDC-UX"

                record = {
                    "department": mapped_department,
                    "source_type": source_type,
                    "project_name": str(proj_name).strip(),
                    "pic": pic,
                    "product_code": prod_code,
                    "bu": mapped_bu,
                    "bg": mapped_bg,
                    "hours": split_hours
                }
                records.append(record)

    # Second pass: Parse the 'Total' sheet to capture leftover hours (e.g. individuals parallel to sub-teams)
    if 'Total' in sheets:
        try:
            df_tot = xl.parse('Total', header=5)
            df_tot.columns = [str(c).strip() for c in df_tot.columns]
            if all(col in df_tot.columns for col in required_cols):
                parent_dept = "MBDC-ID" if source_type == "ID_DETAIL" else ("MBDC-UX" if source_type == "UX_DETAIL" else None)
                if parent_dept:
                    for _, row in df_tot.iterrows():
                        proj_name = row.get('PROJECT NAME')
                        if pd.isna(proj_name) or str(proj_name).strip() == '':
                            continue
                        
                        tot_hours = row.get('Hours')
                        if pd.isna(tot_hours) or tot_hours == 0:
                            continue
                            
                        raw_bu = str(row.get('BU')).strip()
                        pic = str(row.get('PIC')).strip() if not pd.isna(row.get('PIC')) else "Unknown"
                        
                        key = (str(proj_name).strip(), pic, raw_bu)
                        h_tot = float(tot_hours)
                        h_sub = sub_team_hours_map.get(key, 0)
                        
                        # Only inject leftover hours that are in Total but missing from sub-teams
                        if h_tot > h_sub + 0.01:
                            diff_hours = h_tot - h_sub
                            prod_codes = [c.strip().upper() for c in raw_bu.split(',') if c.strip()]
                            if not prod_codes:
                                prod_codes = [raw_bu.upper()]
                            split_hours = diff_hours / len(prod_codes)
                            
                            for prod_code in prod_codes:
                                mapped_bu = "Unknown"
                                mapped_bg = "Unknown"
                                if prod_code in bu_mapping:
                                    mapped_bu = bu_mapping[prod_code]['bu']
                                    mapped_bg = bu_mapping[prod_code]['bg']

                                records.append({
                                    "department": parent_dept,
                                    "source_type": source_type,
                                    "project_name": str(proj_name).strip(),
                                    "pic": pic,
                                    "product_code": prod_code,
                                    "bu": mapped_bu,
                                    "bg": mapped_bg,
                                    "hours": split_hours
                                })
        except Exception:
            pass

    xl.close()
    return {
        "week_id": week_id,
        "date": date_str,
        "working_days": 5, # default
        "records": records
    }
