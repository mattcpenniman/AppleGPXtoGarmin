# Apple Garmin CLI

Convert Apple Health running workouts to Garmin-compatible GPX/TCX files and
download activities from Garmin Connect as TCX through the unofficial API.

## Install

Python 3.12 or newer is required.

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
python -m pip install -e .
```

## Convert Apple Health

Copy `.env.example` to `.env`, adjust paths if needed, then run:

```bash
apple-garmin convert --format tcx
```

CLI arguments override `.env` values. Run `apple-garmin convert --help` for all
options. The existing `python apple_to_garmin_gpx.py` entry point remains
available.

## Config

Edit `.env` if you want different input or output paths:

- `APPLE_EXPORT_DIR`
- `APPLE_EXPORT_XML`
- `WORKOUT_ROUTES_DIR`
- `OUTPUT_DIR`
- `GARMIN_CREATOR`
- `OUTPUT_PREFIX`
- `MODE` (`gpx` or `tcx`)
- `LIMIT` (optional, for example `1`)
- `HEART_RATE_SOURCE` (`record`; `motion_context` is intentionally rejected because it is not BPM)
- `FUZZY_MATCH_SECONDS` (optional small gap fill, for example `10`)
- `DEBUG_XLSX` (`true` to save an Excel debug workbook beside each output)
- `OVERWRITE_EXISTING`

## Export From Garmin

`TSX` is not a Garmin activity format; this command exports Garmin's `TCX`
format. Export the latest activity:

```bash
apple-garmin garmin export-tcx
```

Export specific activities:

```bash
apple-garmin garmin export-tcx --activity-id 123456789 --activity-id 987654321
```

Export up to 20 running activities in a date range:

```bash
apple-garmin garmin export-tcx --start-date 2026-01-01 --end-date 2026-01-31 --activity-type running --limit 20
```

The first run prompts for your Garmin email, password, and MFA code if needed.
OAuth tokens are then stored in `~/.garminconnect`; passwords are not stored by
this project. `GARMIN_EMAIL`, `GARMIN_PASSWORD`, and `GARMINTOKENS` can be
supplied as process environment variables for non-interactive use, but an
interactive password prompt is safer.

Garmin Connect has no supported public activity API. This integration depends on
unofficial web endpoints and may break if Garmin changes them.

## Other Tools

- `python apple_health_explorer.py` opens the local Apple Health explorer.
- `python garmin_batch_import.py` runs the legacy macOS browser import helper.
