# Apple Health Explorer

Run:

```bash
python3 apple_health_explorer.py
```

Then open:

```text
http://127.0.0.1:8765
```

Features:

- browse all `Record` types found in `apple_health_export/export.xml`
- browse all `Workout` activity types found in `apple_health_export/export.xml`
- filter records by start date, end date, source name, and limit
- filter workouts by activity type, start date, end date, source name, and limit
- inspect raw record metadata attached to each row
- inspect raw workout attributes and attached metadata
- cached record-type catalog for faster reloads
- `Update cache` button to rebuild the dropdown catalog from `export.xml`

Notes:

- this is a simple local explorer, not a production app
- queries scan `export.xml` directly, so large filters may take a few seconds
- result limit is capped at `500` per request
