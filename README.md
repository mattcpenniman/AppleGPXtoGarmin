# Apple to Garmin GPX

This script converts Apple Health running route GPX files into Garmin-style GPX files and adds the required:

```xml
<type>running</type>
```

## Files

- `apple_to_garmin_gpx.py` reads `apple_health_export/export.xml`
- It matches only `HKWorkoutActivityTypeRunning` workouts that have a `WorkoutRoute`
- It reads the matching GPX files from `apple_health_export/workout-routes`
- It writes converted files into the folder set in `.env`

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
- `OVERWRITE_EXISTING`

## Run In An IDE

Open `apple_to_garmin_gpx.py` and run it directly.

It uses:

- the local `.env` file for configuration
- `if __name__ == "__main__": main()` so it can run as a normal script
- only the Python standard library, so no package install is required

## Terminal Run

```bash
python3 apple_to_garmin_gpx.py
```

## Output

The script prints a summary including:

- how many running workouts were found in `export.xml`
- how many Garmin GPX files were written
- how many route files were missing
- which output folder was used
