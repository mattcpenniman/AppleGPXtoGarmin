# Garmin Batch Import

This helper opens Garmin Connect import in your normal Chrome session, clicks Garmin's `Browse` control, and selects each batch directly from your source folder without Selenium.

## Install

```bash
python3 garmin_batch_import.py
```

## Behavior

- opens [Garmin Import Data](https://connect.garmin.com/app/import-data)
- lets you log in first if needed
- clicks Garmin's `Browse` control for each batch
- selects the next 10 files directly from your source folder
- can drive the macOS file picker into that folder and select the current batch
- clicks Garmin's `Import Data` button after file selection
- waits only for an initial `ready` confirmation so it does not start before you're logged in

## Config

These values can be added to `.env`:

- `GARMIN_IMPORT_URL`
- `GARMIN_IMPORT_FOLDER`
- `GARMIN_IMPORT_FILE_PATTERN`
- `GARMIN_IMPORT_BATCH_SIZE`
- `GARMIN_IMPORT_STATE_FILE`
- `GARMIN_IMPORT_RESET_PROGRESS`
- `GARMIN_IMPORT_AUTOMATE_DIALOG`
- `GARMIN_IMPORT_BROWSE_CLICK_DELAY`
- `GARMIN_IMPORT_IMPORT_CLICK_DELAY`
- `GARMIN_IMPORT_BATCH_WAIT_SECONDS`
- `IMPORT_DEBUG`

The state file remembers which batch comes next.

If `IMPORT_DEBUG=true`, the script pauses before each major step and asks you to type `go` in the terminal so you can watch the browser flow closely.
