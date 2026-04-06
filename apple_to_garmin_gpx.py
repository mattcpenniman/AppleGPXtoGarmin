from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import re
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET


APPLE_GPX_NS = "http://www.topografix.com/GPX/1/1"
GARMIN_TPX_NS = "http://www.garmin.com/xmlschemas/TrackPointExtension/v1"
GARMIN_GPX_EXT_NS = "http://www.garmin.com/xmlschemas/GpxExtensions/v3"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"


ET.register_namespace("", APPLE_GPX_NS)
ET.register_namespace("xsi", XSI_NS)


@dataclass
class Config:
    apple_export_dir: Path
    export_xml_path: Path
    workout_routes_dir: Path
    output_dir: Path
    garmin_creator: str
    output_prefix: str
    mode: str
    limit: int | None
    heart_rate_source: str
    fuzzy_match_seconds: int
    debug_xlsx: bool
    overwrite_existing: bool


@dataclass
class WorkoutInfo:
    route_relative_path: str | None
    start_date: datetime
    end_date: datetime
    duration_minutes: str | None
    distance: str | None
    distance_unit: str | None
    source_name: str | None
    is_indoor: bool


@dataclass
class SampleInterval:
    start: datetime
    end: datetime
    value: int


@dataclass
class WorkoutMetrics:
    heart_rate: list[SampleInterval]
    cadence: list[SampleInterval]
    source_records: list[dict[str, str]]


def main() -> None:
    project_root = Path(__file__).resolve().parent
    config = load_config(project_root)
    route_map = load_running_workout_routes(config.export_xml_path)
    summary = convert_routes(config, route_map)

    print(f"Running workouts found in export.xml: {len(route_map)}")
    print(f"Converted {config.mode.upper()} files written: {summary['written']}")
    print(f"Skipped existing files: {summary['skipped_existing']}")
    print(f"Missing route files: {summary['missing_routes']}")
    print(f"Output folder: {config.output_dir}")


def load_config(project_root: Path) -> Config:
    env_path = project_root / ".env"
    env_values = parse_env_file(env_path)

    apple_export_dir = resolve_path(
        project_root, env_values.get("APPLE_EXPORT_DIR", "apple_health_export")
    )
    export_xml_path = resolve_path(
        project_root, env_values.get("APPLE_EXPORT_XML", str(apple_export_dir / "export.xml"))
    )
    workout_routes_dir = resolve_path(
        project_root,
        env_values.get("WORKOUT_ROUTES_DIR", str(apple_export_dir / "workout-routes")),
    )
    output_dir = resolve_path(project_root, env_values.get("OUTPUT_DIR", "garmin_gpx_output"))

    return Config(
        apple_export_dir=apple_export_dir,
        export_xml_path=export_xml_path,
        workout_routes_dir=workout_routes_dir,
        output_dir=output_dir,
        garmin_creator=env_values.get("GARMIN_CREATOR", "Garmin Connect"),
        output_prefix=env_values.get("OUTPUT_PREFIX", "activity"),
        mode=parse_mode(env_values.get("MODE", "gpx")),
        limit=parse_optional_int(env_values.get("LIMIT")),
        heart_rate_source=parse_heart_rate_source(
            env_values.get("HEART_RATE_SOURCE", "record")
        ),
        fuzzy_match_seconds=parse_non_negative_int(env_values.get("FUZZY_MATCH_SECONDS", "0")),
        debug_xlsx=parse_bool(env_values.get("DEBUG_XLSX", "false")),
        overwrite_existing=parse_bool(env_values.get("OVERWRITE_EXISTING", "true")),
    )


def parse_env_file(env_path: Path) -> dict[str, str]:
    if not env_path.exists():
        raise FileNotFoundError(f"Missing .env file: {env_path}")

    values: dict[str, str] = {}
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()

        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]

        values[key] = value

    return values


def parse_bool(value: str) -> bool:
    return value.strip().lower() in {"1", "true", "yes", "on"}


def parse_mode(value: str) -> str:
    mode = value.strip().lower()
    if mode not in {"gpx", "tcx"}:
        raise ValueError("MODE must be either 'gpx' or 'tcx'")
    return mode


def parse_optional_int(value: str | None) -> int | None:
    if value is None:
        return None

    stripped = value.strip()
    if not stripped:
        return None

    limit = int(stripped)
    if limit <= 0:
        raise ValueError("LIMIT must be a positive integer when provided")
    return limit


def parse_non_negative_int(value: str) -> int:
    parsed = int(value.strip())
    if parsed < 0:
        raise ValueError("FUZZY_MATCH_SECONDS must be zero or greater")
    return parsed


def parse_heart_rate_source(value: str) -> str:
    source = value.strip().lower()
    if source not in {"record", "motion_context"}:
        raise ValueError("HEART_RATE_SOURCE must be 'record' or 'motion_context'")
    return source


def resolve_path(project_root: Path, raw_path: str) -> Path:
    path = Path(raw_path).expanduser()
    if not path.is_absolute():
        path = project_root / path
    return path.resolve()


def load_running_workout_routes(export_xml_path: Path) -> dict[str, WorkoutInfo]:
    if not export_xml_path.exists():
        raise FileNotFoundError(f"Apple export.xml not found: {export_xml_path}")

    route_map: dict[str, WorkoutInfo] = {}

    for _, elem in ET.iterparse(export_xml_path, events=("end",)):
        if elem.tag != "Workout":
            continue

        if elem.attrib.get("workoutActivityType") != "HKWorkoutActivityTypeRunning":
            elem.clear()
            continue

        route_relative_path: str | None = None
        is_indoor = False
        distance = elem.attrib.get("totalDistance")
        distance_unit = elem.attrib.get("totalDistanceUnit")
        for child in elem:
            if (
                child.tag == "MetadataEntry"
                and child.attrib.get("key") == "HKIndoorWorkout"
                and child.attrib.get("value") == "1"
            ):
                is_indoor = True
            if child.tag == "WorkoutStatistics":
                stat_type = child.attrib.get("type")
                if (
                    not distance
                    and stat_type == "HKQuantityTypeIdentifierDistanceWalkingRunning"
                ):
                    distance = child.attrib.get("sum")
                    distance_unit = child.attrib.get("unit")
            if child.tag != "WorkoutRoute":
                continue
            for route_child in child:
                if route_child.tag == "FileReference":
                    route_relative_path = route_child.attrib.get("path")
                    break
            if route_relative_path:
                break

        workout_key = normalize_route_path(route_relative_path) if route_relative_path else build_workout_key(elem.attrib["startDate"])

        if route_relative_path or is_indoor:
            route_map[workout_key] = WorkoutInfo(
                route_relative_path=normalize_route_path(route_relative_path) if route_relative_path else None,
                start_date=parse_apple_datetime(elem.attrib["startDate"]),
                end_date=parse_apple_datetime(elem.attrib["endDate"]),
                duration_minutes=elem.attrib.get("duration"),
                distance=distance,
                distance_unit=distance_unit,
                source_name=elem.attrib.get("sourceName"),
                is_indoor=is_indoor,
            )

        elem.clear()

    return route_map


def normalize_route_path(route_path: str) -> str:
    cleaned = route_path.replace("\\", "/").strip()
    return cleaned.lstrip("/")


def build_workout_key(start_date: str) -> str:
    return f"__workout__/{start_date}"


def parse_apple_datetime(value: str) -> datetime:
    return datetime.strptime(value, "%Y-%m-%d %H:%M:%S %z")


def convert_routes(config: Config, route_map: dict[str, WorkoutInfo]) -> dict[str, int]:
    config.output_dir.mkdir(parents=True, exist_ok=True)

    summary = {"written": 0, "skipped_existing": 0, "missing_routes": 0}

    items = sorted(route_map.items())
    if config.limit is not None:
        items = items[: config.limit]

    metrics_by_route = load_workout_metrics(
        config.export_xml_path,
        items,
        config.heart_rate_source,
    )
    export_catalogs = load_export_catalogs(config.export_xml_path) if config.debug_xlsx else None

    for route_key, workout in items:
        route_file: Path | None = None
        if workout.route_relative_path:
            route_file = config.apple_export_dir / workout.route_relative_path
            if not route_file.exists():
                route_file = config.workout_routes_dir / Path(workout.route_relative_path).name

            if not route_file.exists():
                print(f"Missing route file for running workout: {route_key}")
                summary["missing_routes"] += 1
                continue

        if config.mode == "gpx" and route_file is None:
            print(f"Skipping workout without route in GPX mode: {workout.start_date}")
            continue

        output_path = config.output_dir / build_output_filename(
            config.output_prefix,
            workout,
            config.mode,
        )
        if output_path.exists() and not config.overwrite_existing:
            summary["skipped_existing"] += 1
            continue

        if config.mode == "gpx":
            output_tree = build_garmin_gpx_tree(
                apple_route_path=route_file,
                workout=workout,
                metrics=metrics_by_route.get(route_key, WorkoutMetrics([], [], [])),
                fuzzy_match_seconds=config.fuzzy_match_seconds,
                garmin_creator=config.garmin_creator,
            )
        else:
            output_tree = build_garmin_tcx_tree(
                apple_route_path=route_file,
                workout=workout,
                metrics=metrics_by_route.get(route_key, WorkoutMetrics([], [], [])),
                fuzzy_match_seconds=config.fuzzy_match_seconds,
            )

        indent_xml(output_tree)
        output_tree.write(output_path, encoding="utf-8", xml_declaration=True)
        if config.debug_xlsx:
            write_debug_workbook(
                output_path=output_path,
                export_catalogs=export_catalogs,
                apple_route_path=route_file,
                workout=workout,
                metrics=metrics_by_route.get(route_key, WorkoutMetrics([], [], [])),
                fuzzy_match_seconds=config.fuzzy_match_seconds,
            )
        summary["written"] += 1

    return summary


def build_output_filename(output_prefix: str, workout: WorkoutInfo, mode: str) -> str:
    timestamp = workout.start_date.strftime("%Y%m%d_%H%M%S")
    return f"{output_prefix}_{timestamp}.{mode}"


def build_garmin_gpx_tree(
    apple_route_path: Path | None,
    workout: WorkoutInfo,
    metrics: WorkoutMetrics,
    fuzzy_match_seconds: int,
    garmin_creator: str,
) -> ET.ElementTree:
    if apple_route_path is None:
        raise ValueError("GPX export requires a route file")
    apple_tree = ET.parse(apple_route_path)
    apple_root = apple_tree.getroot()

    garmin_root = ET.Element(
        f"{{{APPLE_GPX_NS}}}gpx",
        {
            "creator": garmin_creator,
            "version": "1.1",
            "xmlns:ns2": GARMIN_GPX_EXT_NS,
            "xmlns:ns3": GARMIN_TPX_NS,
            f"{{{XSI_NS}}}schemaLocation": (
                "http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/11.xsd"
            ),
        },
    )

    metadata = ET.SubElement(garmin_root, f"{{{APPLE_GPX_NS}}}metadata")
    link = ET.SubElement(metadata, f"{{{APPLE_GPX_NS}}}link", {"href": "connect.garmin.com"})
    ET.SubElement(link, f"{{{APPLE_GPX_NS}}}text").text = "Garmin Connect"
    ET.SubElement(metadata, f"{{{APPLE_GPX_NS}}}time").text = format_garmin_time(workout.start_date)

    trk = ET.SubElement(garmin_root, f"{{{APPLE_GPX_NS}}}trk")
    ET.SubElement(trk, f"{{{APPLE_GPX_NS}}}name").text = build_track_name(workout)
    ET.SubElement(trk, f"{{{APPLE_GPX_NS}}}type").text = "running"
    trkseg = ET.SubElement(trk, f"{{{APPLE_GPX_NS}}}trkseg")

    apple_trkpts = apple_root.findall(f".//{{{APPLE_GPX_NS}}}trkpt")
    if not apple_trkpts:
        raise ValueError(f"No track points found in {apple_route_path}")

    hr_cursor = 0
    cad_cursor = 0
    for apple_trkpt in apple_trkpts:
        garmin_trkpt = ET.SubElement(
            trkseg,
            f"{{{APPLE_GPX_NS}}}trkpt",
            {
                "lat": apple_trkpt.attrib["lat"],
                "lon": apple_trkpt.attrib["lon"],
            },
        )

        apple_ele = apple_trkpt.find(f"{{{APPLE_GPX_NS}}}ele")
        apple_time = apple_trkpt.find(f"{{{APPLE_GPX_NS}}}time")

        if apple_ele is not None and apple_ele.text:
            ET.SubElement(garmin_trkpt, f"{{{APPLE_GPX_NS}}}ele").text = apple_ele.text
        if apple_time is not None and apple_time.text:
            point_time = datetime.fromisoformat(apple_time.text.replace("Z", "+00:00"))
            ET.SubElement(garmin_trkpt, f"{{{APPLE_GPX_NS}}}time").text = normalize_utc_text(apple_time.text)
            heart_rate, hr_cursor = sample_value_for_time(
                metrics.heart_rate,
                point_time,
                hr_cursor,
                fuzzy_match_seconds,
            )
            cadence, cad_cursor = sample_value_for_time(
                metrics.cadence,
                point_time,
                cad_cursor,
                fuzzy_match_seconds,
            )
            if heart_rate is not None or cadence is not None:
                extensions = ET.SubElement(garmin_trkpt, f"{{{APPLE_GPX_NS}}}extensions")
                tpx = ET.SubElement(extensions, f"{{{GARMIN_TPX_NS}}}TrackPointExtension")
                if heart_rate is not None:
                    ET.SubElement(tpx, f"{{{GARMIN_TPX_NS}}}hr").text = str(heart_rate)
                if cadence is not None:
                    ET.SubElement(tpx, f"{{{GARMIN_TPX_NS}}}cad").text = str(cadence)

    return ET.ElementTree(garmin_root)


def build_garmin_tcx_tree(
    apple_route_path: Path | None,
    workout: WorkoutInfo,
    metrics: WorkoutMetrics,
    fuzzy_match_seconds: int,
) -> ET.ElementTree:
    apple_trkpts: list[ET.Element] = []
    if apple_route_path is not None:
        apple_tree = ET.parse(apple_route_path)
        apple_root = apple_tree.getroot()
        apple_trkpts = apple_root.findall(f".//{{{APPLE_GPX_NS}}}trkpt")

    tcx_ns = "http://www.garmin.com/xmlschemas/TrainingCenterDatabase/v2"
    xsi_schema = (
        "http://www.garmin.com/xmlschemas/TrainingCenterDatabase/v2 "
        "http://www.garmin.com/xmlschemas/TrainingCenterDatabasev2.xsd"
    )

    ET.register_namespace("", tcx_ns)

    root = ET.Element(
        f"{{{tcx_ns}}}TrainingCenterDatabase",
        {f"{{{XSI_NS}}}schemaLocation": xsi_schema},
    )
    activities = ET.SubElement(root, f"{{{tcx_ns}}}Activities")
    activity = ET.SubElement(activities, f"{{{tcx_ns}}}Activity", {"Sport": "Running"})
    ET.SubElement(activity, f"{{{tcx_ns}}}Id").text = format_tcx_time(workout.start_date)

    lap = ET.SubElement(
        activity,
        f"{{{tcx_ns}}}Lap",
        {"StartTime": format_tcx_time(workout.start_date)},
    )
    ET.SubElement(lap, f"{{{tcx_ns}}}TotalTimeSeconds").text = format_seconds(workout)
    ET.SubElement(lap, f"{{{tcx_ns}}}DistanceMeters").text = format_distance_meters(workout)
    ET.SubElement(lap, f"{{{tcx_ns}}}Calories").text = "0"
    ET.SubElement(lap, f"{{{tcx_ns}}}Intensity").text = "Active"
    ET.SubElement(lap, f"{{{tcx_ns}}}TriggerMethod").text = "Manual"

    if apple_trkpts:
        track = ET.SubElement(lap, f"{{{tcx_ns}}}Track")
        hr_cursor = 0
        cad_cursor = 0
        for apple_trkpt in apple_trkpts:
            trackpoint = ET.SubElement(track, f"{{{tcx_ns}}}Trackpoint")
            apple_time = apple_trkpt.find(f"{{{APPLE_GPX_NS}}}time")
            apple_ele = apple_trkpt.find(f"{{{APPLE_GPX_NS}}}ele")

            if apple_time is not None and apple_time.text:
                point_time = datetime.fromisoformat(apple_time.text.replace("Z", "+00:00"))
                ET.SubElement(trackpoint, f"{{{tcx_ns}}}Time").text = normalize_utc_text(apple_time.text)
                heart_rate, hr_cursor = sample_value_for_time(
                    metrics.heart_rate,
                    point_time,
                    hr_cursor,
                    fuzzy_match_seconds,
                )
                cadence, cad_cursor = sample_value_for_time(
                    metrics.cadence,
                    point_time,
                    cad_cursor,
                    fuzzy_match_seconds,
                )
                if heart_rate is not None:
                    heart_rate_bpm = ET.SubElement(trackpoint, f"{{{tcx_ns}}}HeartRateBpm")
                    ET.SubElement(heart_rate_bpm, f"{{{tcx_ns}}}Value").text = str(heart_rate)
                if cadence is not None:
                    ET.SubElement(trackpoint, f"{{{tcx_ns}}}Cadence").text = str(cadence)

            position = ET.SubElement(trackpoint, f"{{{tcx_ns}}}Position")
            ET.SubElement(position, f"{{{tcx_ns}}}LatitudeDegrees").text = apple_trkpt.attrib["lat"]
            ET.SubElement(position, f"{{{tcx_ns}}}LongitudeDegrees").text = apple_trkpt.attrib["lon"]

            if apple_ele is not None and apple_ele.text:
                ET.SubElement(trackpoint, f"{{{tcx_ns}}}AltitudeMeters").text = apple_ele.text

    return ET.ElementTree(root)


def load_workout_metrics(
    export_xml_path: Path,
    workout_items: list[tuple[str, WorkoutInfo]],
    heart_rate_source: str,
) -> dict[str, WorkoutMetrics]:
    metrics_by_route = {
        route_key: WorkoutMetrics(heart_rate=[], cadence=[], source_records=[])
        for route_key, _ in workout_items
    }
    if not workout_items:
        return metrics_by_route

    sorted_items = sorted(workout_items, key=lambda item: item[1].start_date)
    starts = [workout.start_date for _, workout in sorted_items]

    for _, elem in ET.iterparse(export_xml_path, events=("end",)):
        if elem.tag != "Record":
            continue

        record_type = elem.attrib.get("type")
        if record_type not in {
            "HKQuantityTypeIdentifierHeartRate",
            "HKQuantityTypeIdentifierStepCount",
        }:
            elem.clear()
            continue

        try:
            sample_start = parse_apple_datetime(elem.attrib["startDate"])
            sample_end = parse_apple_datetime(elem.attrib["endDate"])
        except (KeyError, ValueError):
            elem.clear()
            continue

        value_text = elem.attrib.get("value")
        if not value_text:
            elem.clear()
            continue

        source_name = elem.attrib.get("sourceName", "").replace("\xa0", " ")
        watch_source = "Apple Watch" in source_name
        if not watch_source:
            elem.clear()
            continue

        if record_type == "HKQuantityTypeIdentifierHeartRate" and heart_rate_source == "motion_context":
            raise ValueError(
                "HEART_RATE_SOURCE=motion_context is not supported for Garmin export because "
                "HKMetadataKeyHeartRateMotionContext contains categorical values like 0/1/2, not BPM."
            )

        sample_value = convert_record_to_metric_value(
            record_type=record_type,
            value_text=value_text,
            sample_start=sample_start,
            sample_end=sample_end,
        )
        if sample_value is None:
            elem.clear()
            continue

        candidate_index = bisect_right(starts, sample_end)
        for index in range(candidate_index - 1, -1, -1):
            route_key, workout = sorted_items[index]
            if workout.end_date < sample_start:
                break
            if workout.start_date <= sample_end and workout.end_date >= sample_start:
                interval = SampleInterval(sample_start, sample_end, sample_value)
                if record_type == "HKQuantityTypeIdentifierHeartRate":
                    metrics_by_route[route_key].heart_rate.append(interval)
                else:
                    metrics_by_route[route_key].cadence.append(interval)
                metrics_by_route[route_key].source_records.append(
                    {
                        "record_type": record_type,
                        "start": sample_start.isoformat(),
                        "end": sample_end.isoformat(),
                        "value": value_text,
                        "unit": elem.attrib.get("unit", ""),
                        "source_name": source_name,
                    }
                )
        elem.clear()

    for metrics in metrics_by_route.values():
        metrics.heart_rate.sort(key=lambda interval: interval.start)
        metrics.cadence.sort(key=lambda interval: interval.start)

    return metrics_by_route


def convert_record_to_metric_value(
    record_type: str,
    value_text: str,
    sample_start: datetime,
    sample_end: datetime,
) -> int | None:
    try:
        value = float(value_text)
    except ValueError:
        return None

    if record_type == "HKQuantityTypeIdentifierHeartRate":
        return round(value)

    duration_seconds = max((sample_end - sample_start).total_seconds(), 0)
    if duration_seconds <= 0:
        return None

    steps_per_minute = (value * 60) / duration_seconds
    return round(steps_per_minute)


def sample_value_for_time(
    intervals: list[SampleInterval],
    point_time: datetime,
    cursor: int,
    fuzzy_match_seconds: int,
) -> tuple[int | None, int]:
    while cursor < len(intervals) and intervals[cursor].end < point_time:
        cursor += 1

    if cursor < len(intervals):
        interval = intervals[cursor]
        if interval.start <= point_time <= interval.end:
            return interval.value, cursor

    if fuzzy_match_seconds <= 0:
        return None, cursor

    candidates: list[tuple[float, int, int]] = []
    if cursor < len(intervals):
        next_interval = intervals[cursor]
        candidates.append(
            (
                boundary_distance_seconds(next_interval, point_time),
                next_interval.value,
                cursor,
            )
        )
    if cursor > 0:
        previous_interval = intervals[cursor - 1]
        candidates.append(
            (
                boundary_distance_seconds(previous_interval, point_time),
                previous_interval.value,
                cursor - 1,
            )
        )

    if candidates:
        delta, value, matched_cursor = min(candidates, key=lambda item: item[0])
        if delta <= fuzzy_match_seconds:
            return value, matched_cursor

    return None, cursor


def boundary_distance_seconds(interval: SampleInterval, point_time: datetime) -> float:
    if interval.start <= point_time <= interval.end:
        return 0.0
    if point_time < interval.start:
        return (interval.start - point_time).total_seconds()
    return (point_time - interval.end).total_seconds()


def build_track_name(workout: WorkoutInfo) -> str:
    local_time = workout.start_date.strftime("%Y-%m-%d %I:%M %p").lstrip("0")
    details: list[str] = []

    if workout.distance and workout.distance_unit:
        details.append(f"{trim_number(workout.distance)} {workout.distance_unit}")
    if workout.duration_minutes:
        details.append(f"{trim_number(workout.duration_minutes)} min")

    suffix = f" ({', '.join(details)})" if details else ""
    return f"Run {local_time}{suffix}"


def trim_number(value: str) -> str:
    trimmed = re.sub(r"(\.\d*?[1-9])0+$", r"\1", value)
    trimmed = re.sub(r"\.0+$", "", trimmed)
    return trimmed


def normalize_utc_text(value: str) -> str:
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return value
    return format_garmin_time(parsed)


def format_garmin_time(value: datetime) -> str:
    utc_value = value.astimezone(timezone.utc)
    return utc_value.strftime("%Y-%m-%dT%H:%M:%S.000Z")


def format_tcx_time(value: datetime) -> str:
    utc_value = value.astimezone(timezone.utc)
    return utc_value.strftime("%Y-%m-%dT%H:%M:%SZ")


def format_seconds(workout: WorkoutInfo) -> str:
    if workout.duration_minutes is None:
        seconds = max((workout.end_date - workout.start_date).total_seconds(), 0)
    else:
        seconds = float(workout.duration_minutes) * 60
    return f"{seconds:.3f}"


def format_distance_meters(workout: WorkoutInfo) -> str:
    if not workout.distance or not workout.distance_unit:
        return "0"

    distance = float(workout.distance)
    unit = workout.distance_unit.lower()
    if unit == "mi":
        meters = distance * 1609.344
    elif unit == "km":
        meters = distance * 1000
    elif unit == "m":
        meters = distance
    else:
        meters = distance
    return f"{meters:.3f}"


def write_debug_workbook(
    output_path: Path,
    export_catalogs: tuple[list[list[str]], list[list[str]]] | None,
    apple_route_path: Path | None,
    workout: WorkoutInfo,
    metrics: WorkoutMetrics,
    fuzzy_match_seconds: int,
) -> None:
    workbook_path = output_path.with_suffix(".debug.xlsx")
    sheets = {
        "route_points": build_debug_points_rows(
            apple_route_path=apple_route_path,
            metrics=metrics,
            fuzzy_match_seconds=fuzzy_match_seconds,
        ),
        "source_records": build_debug_records_rows(workout, metrics),
    }
    if export_catalogs is not None:
        metadata_rows, record_rows = export_catalogs
        sheets["metadata_keys"] = metadata_rows
        sheets["record_types"] = record_rows

    create_xlsx(
        workbook_path,
        sheets,
    )


def build_debug_points_rows(
    apple_route_path: Path | None,
    metrics: WorkoutMetrics,
    fuzzy_match_seconds: int,
) -> list[list[str]]:
    if apple_route_path is None:
        return [["point_time", "lat", "lon", "ele", "matched_hr", "matched_cadence"]]
    apple_tree = ET.parse(apple_route_path)
    apple_root = apple_tree.getroot()
    apple_trkpts = apple_root.findall(f".//{{{APPLE_GPX_NS}}}trkpt")

    rows = [[
        "point_time",
        "lat",
        "lon",
        "ele",
        "matched_hr",
        "matched_cadence",
    ]]
    hr_cursor = 0
    cad_cursor = 0

    for apple_trkpt in apple_trkpts:
        apple_time = apple_trkpt.find(f"{{{APPLE_GPX_NS}}}time")
        apple_ele = apple_trkpt.find(f"{{{APPLE_GPX_NS}}}ele")
        point_time = None
        heart_rate = None
        cadence = None

        if apple_time is not None and apple_time.text:
            point_time = datetime.fromisoformat(apple_time.text.replace("Z", "+00:00"))
            heart_rate, hr_cursor = sample_value_for_time(
                metrics.heart_rate,
                point_time,
                hr_cursor,
                fuzzy_match_seconds,
            )
            cadence, cad_cursor = sample_value_for_time(
                metrics.cadence,
                point_time,
                cad_cursor,
                fuzzy_match_seconds,
            )

        rows.append([
            apple_time.text if apple_time is not None and apple_time.text else "",
            apple_trkpt.attrib.get("lat", ""),
            apple_trkpt.attrib.get("lon", ""),
            apple_ele.text if apple_ele is not None and apple_ele.text else "",
            "" if heart_rate is None else str(heart_rate),
            "" if cadence is None else str(cadence),
        ])

    return rows


def build_debug_records_rows(
    workout: WorkoutInfo,
    metrics: WorkoutMetrics,
) -> list[list[str]]:
    rows = [[
        "run_start",
        "run_end",
        "record_type",
        "record_start",
        "record_end",
        "value",
        "unit",
        "source_name",
    ]]
    for record in sorted(metrics.source_records, key=lambda item: (item["record_type"], item["start"])):
        rows.append([
            workout.start_date.isoformat(),
            workout.end_date.isoformat(),
            record["record_type"],
            record["start"],
            record["end"],
            record["value"],
            record["unit"],
            record["source_name"],
        ])
    return rows


def load_export_catalogs(export_xml_path: Path) -> tuple[list[list[str]], list[list[str]]]:
    metadata_counts: dict[str, int] = {}
    metadata_examples: dict[str, str] = {}
    record_counts: dict[str, int] = {}
    record_units: dict[str, str] = {}
    record_sources: dict[str, str] = {}

    for _, elem in ET.iterparse(export_xml_path, events=("end",)):
        if elem.tag == "MetadataEntry":
            key = elem.attrib.get("key")
            if key:
                metadata_counts[key] = metadata_counts.get(key, 0) + 1
                metadata_examples.setdefault(key, elem.attrib.get("value", ""))
        elif elem.tag == "Record":
            record_type = elem.attrib.get("type")
            if record_type:
                record_counts[record_type] = record_counts.get(record_type, 0) + 1
                record_units.setdefault(record_type, elem.attrib.get("unit", ""))
                record_sources.setdefault(
                    record_type,
                    elem.attrib.get("sourceName", "").replace("\xa0", " "),
                )
        elem.clear()

    metadata_rows = [[
        "metadata_key",
        "count",
        "example_value",
    ]]
    for key in sorted(metadata_counts):
        metadata_rows.append([
            key,
            str(metadata_counts[key]),
            metadata_examples.get(key, ""),
        ])

    record_rows = [[
        "record_type",
        "count",
        "unit",
        "example_source",
    ]]
    for record_type in sorted(record_counts):
        record_rows.append([
            record_type,
            str(record_counts[record_type]),
            record_units.get(record_type, ""),
            record_sources.get(record_type, ""),
        ])

    return metadata_rows, record_rows


def create_xlsx(path: Path, sheets: dict[str, list[list[str]]]) -> None:
    sheet_names = list(sheets.keys())
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", build_content_types_xml(len(sheet_names)))
        zf.writestr("_rels/.rels", ROOT_RELS_XML)
        zf.writestr("xl/workbook.xml", build_workbook_xml(sheet_names))
        zf.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels_xml(len(sheet_names)))
        for index, sheet_name in enumerate(sheet_names, start=1):
            zf.writestr(
                f"xl/worksheets/sheet{index}.xml",
                build_sheet_xml(sheets[sheet_name]),
            )


def build_content_types_xml(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides)
        + "</Types>"
    )


ROOT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="xl/workbook.xml"/>'
    '</Relationships>'
)


def build_workbook_xml(sheet_names: list[str]) -> str:
    sheets_xml = []
    for index, name in enumerate(sheet_names, start=1):
        sheets_xml.append(
            f'<sheet name="{xml_escape(name[:31])}" sheetId="{index}" r:id="rId{index}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{''.join(sheets_xml)}</sheets>"
        "</workbook>"
    )


def build_workbook_rels_xml(sheet_count: int) -> str:
    rels_xml = []
    for index in range(1, sheet_count + 1):
        rels_xml.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels_xml)
        + "</Relationships>"
    )


def build_sheet_xml(rows: list[list[str]]) -> str:
    row_xml = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            cell_ref = f"{excel_column_name(col_index)}{row_index}"
            cells.append(
                f'<c r="{cell_ref}" t="inlineStr"><is><t>{xml_escape(value)}</t></is></c>'
            )
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(row_xml)}</sheetData>"
        "</worksheet>"
    )


def excel_column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def xml_escape(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def indent_xml(tree: ET.ElementTree) -> None:
    try:
        ET.indent(tree, space="  ")
    except AttributeError:
        pass


if __name__ == "__main__":
    main()
