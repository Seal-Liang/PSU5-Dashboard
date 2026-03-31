# Project Constitution

## Data Schemas
The system relies on JSON structures to store and process the uploaded Excel data, BU mappings, and system settings.

```json
{
  "Settings": {
    "headcount_logs": [
      {
        "department": "MBDC-ADS",
        "initial": 10,
        "logs": [
          { "date": "2025-08-01", "type": "到職", "delta": 1, "name": "Seal Liang", "note": "" }
        ]
      }
    ],
    "holidays": ["2025-01-01", "2025-02-28"]
  },
  "BUMapping": [
    {
      "product_code": "BZHY",
      "bu": "BU-01",
      "bg": "BG-01",
      "product_name": "Product X"
    }
  ],
  "WeeklyData": [
    {
      "week_id": "2025WK31",
      "date": "2025-07-28",
      "working_days": 5,
      "records": [
        {
          "department": "MBDC-ADS",
          "project_name": "Ahbpavpxhb_Vpulcrp79",
          "pic": "Hayden Su",
          "product_code": "BZHY",
          "hours": 15
        }
      ]
    }
  ]
}
```

## Behavioral Rules
- **Tone & Style:** Clean, professional, minimal, data-analysis focused with clear charts.
- **Data Persistence:** Save previously uploaded data locally (JSON/SQLite) so users only need to upload new files.
- **Dynamic Calculation:** Utilization rate must be calculated dynamically based on configurable department headcount and weekly working days.
- **Excel Parsing Invariant:** Weekly Excel reports have headers starting at Row 5 (0-indexed). The "BU" column in the Excel maps to "Product Code" in the BU Mapping file.

## Architectural Invariants
- 3-Layer Architecture (A.N.T.)
- Business logic must be deterministic
- No tools built until Schema confirmed

## Maintenance Log
- **Local Deployment:** Run `python server.py` and open `http://localhost:5000` to interact with the dashboard.
- **Python Dependencies:** Requires `Flask`, `pandas`, `openpyxl`, and `python-dotenv`.
- **Data Shape:** Changing the Excel source format (e.g., migrating from Row 5 headers) requires a Layer 1 SOP review (`architecture/1_parser_sop.md`) followed by a Layer 3 Tool update (`tools/excel_parser.py`).
