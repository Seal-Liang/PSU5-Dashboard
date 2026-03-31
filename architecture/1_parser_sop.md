# Excel Parser SOP (Layer 1)

## Goal
Extract weekly project hours and BU mapping deterministically from user-uploaded Excel files.

## Invariants
- The JSON Data Schema relies on headers located strictly at Row 5 (`header=5` in `pandas`).
- The `BU` column in the weekly report maps exactly to `Product Code` in the BU classification sheet.

## Tool Logic (`tools/excel_parser.py`)
1. Receive paths to `WeeklyReport.xlsx` and `BUMapping.xlsx`.
2. Parse `BUMapping.xlsx`: Create a lookup dictionary `Product Code` -> `{BU, BG, Product Name}`.
3. Parse `WeeklyReport.xlsx`:
   - Extract `Week ID` and `Date` from the first sheet, Row 0, Column 0.
   - For every sheet EXCEPT "Total" and "AMS" (or equivalent skip tabs):
     - Read data starting at Row 5.
     - Filter out rows where `PROJECT NAME` is empty/NaN, or where `Hours` is 0 or NaN.
     - Extract `Hours`, `PIC`, and `Product Code` (From `BU` column).
     - Construct a list of records for this department.
4. Return a normalized JSON shape conforming to the `WeeklyData` schema in `gemini.md`.

## Edge Cases
- Missing Row 5 headers (Throw explicit error).
- Projects with 0 hours (Skip them to save space).
- Unknown `BU` codes (Keep as-is).
