# Navigation SOP (Layer 2)

## Goal
Route file uploads and UI chart data requests deterministically.

## Logic (`server.py`)
1. Frontend calls `POST /api/upload` with FormData containing `weekly_report` and `bu_mapping`.
2. `server.py` temporarily saves files to `.tmp/`.
3. Validation: Verify files have `.xlsx` extension.
4. Route to Layer 3 Tools:
   - Call `parse_bu_mapping(bu_filepath)`
   - Call `parse_weekly_report(weekly_filepath, mapping)`
5. Pass resulting objects to `data_store.py`:
   - `save_bu_mapping(...)`
   - `save_weekly_data(...)`
6. Return `200 {"status": "success"}` or `400 / 500 {"error": "message"}` upon deterministic failure.
7. Frontend calls `GET /api/data` -> Server serves `data.json` state.
