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

def parse_weekly_report(filepath, bu_mapping):
    """
    Parses the Weekly Data Excel file.
    """
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
                
            prod_code = str(row.get('BU')).strip().upper()
            
            # Use mapping if available. The Excel 'BU' column is actually the Product Code
            mapped_bu = "Unknown"
            mapped_bg = "Unknown"
            
            if prod_code in bu_mapping:
                mapped_bu = bu_mapping[prod_code]['bu']
                mapped_bg = bu_mapping[prod_code]['bg']

            record = {
                "department": sheet,
                "project_name": str(proj_name).strip(),
                "pic": str(row.get('PIC')).strip() if not pd.isna(row.get('PIC')) else "Unknown",
                "product_code": prod_code,
                "bu": mapped_bu,
                "bg": mapped_bg,
                "hours": float(hours)
            }
            records.append(record)

    xl.close()
    return {
        "week_id": week_id,
        "date": date_str,
        "working_days": 5, # default
        "records": records
    }
