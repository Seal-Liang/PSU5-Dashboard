# Data Storage SOP (Layer 1)

## Goal
Persist parsed JSON data to `data.json` safely.

## Invariants
- Data is stored locally in `data.json` in the root folder.
- The state must strictly follow the Schema defined in `gemini.md`.

## Tool Logic (`tools/data_store.py`)
1. Read existing `data.json`. Initialize with empty lists if missing.
2. `save_weekly_data(parsed_data)`:
   - Check if the uploaded `week_id` already exists in `WeeklyData`.
   - If it exists, overwrite the week to allow clean re-uploads.
   - Otherwise, append to `WeeklyData`.
3. `save_bu_mapping(parsed_mapping)`:
   - Overwrite the `BUMapping` array entirely.
4. `update_settings(updates)`:
   - Update department headcounts or holidays incrementally.

## Edge Cases
- File locks or IO errors: Handle gracefully and return JSON error.
- Corrupted `data.json`: Backup existing to `.tmp/` and re-initialize.
