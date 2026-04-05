from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import re
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
    overwrite_existing: bool


@dataclass
class WorkoutInfo:
    route_relative_path: str
    start_date: datetime
    end_date: datetime
    duration_minutes: str | None
    distance: str | None
    distance_unit: str | None
    source_name: str | None


def main() -> None:
    project_root = Path(__file__).resolve().parent
    config = load_config(project_root)
    route_map = load_running_workout_routes(config.export_xml_path)
    summary = convert_routes(config, route_map)

    print(f"Running workouts found in export.xml: {len(route_map)}")
    print(f"Converted GPX files written: {summary['written']}")
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
        distance = elem.attrib.get("totalDistance")
        distance_unit = elem.attrib.get("totalDistanceUnit")
        for child in elem:
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

        if route_relative_path:
            route_map[normalize_route_path(route_relative_path)] = WorkoutInfo(
                route_relative_path=normalize_route_path(route_relative_path),
                start_date=parse_apple_datetime(elem.attrib["startDate"]),
                end_date=parse_apple_datetime(elem.attrib["endDate"]),
                duration_minutes=elem.attrib.get("duration"),
                distance=distance,
                distance_unit=distance_unit,
                source_name=elem.attrib.get("sourceName"),
            )

        elem.clear()

    return route_map


def normalize_route_path(route_path: str) -> str:
    cleaned = route_path.replace("\\", "/").strip()
    return cleaned.lstrip("/")


def parse_apple_datetime(value: str) -> datetime:
    return datetime.strptime(value, "%Y-%m-%d %H:%M:%S %z")


def convert_routes(config: Config, route_map: dict[str, WorkoutInfo]) -> dict[str, int]:
    config.output_dir.mkdir(parents=True, exist_ok=True)

    summary = {"written": 0, "skipped_existing": 0, "missing_routes": 0}

    for route_key, workout in sorted(route_map.items()):
        route_file = config.apple_export_dir / route_key
        if not route_file.exists():
            route_file = config.workout_routes_dir / Path(route_key).name

        if not route_file.exists():
            print(f"Missing route file for running workout: {route_key}")
            summary["missing_routes"] += 1
            continue

        output_path = config.output_dir / build_output_filename(config.output_prefix, workout)
        if output_path.exists() and not config.overwrite_existing:
            summary["skipped_existing"] += 1
            continue

        garmin_tree = build_garmin_tree(
            apple_route_path=route_file,
            workout=workout,
            garmin_creator=config.garmin_creator,
        )
        indent_xml(garmin_tree)
        garmin_tree.write(output_path, encoding="utf-8", xml_declaration=True)
        summary["written"] += 1

    return summary


def build_output_filename(output_prefix: str, workout: WorkoutInfo) -> str:
    timestamp = workout.start_date.strftime("%Y%m%d_%H%M%S")
    return f"{output_prefix}_{timestamp}.gpx"


def build_garmin_tree(apple_route_path: Path, workout: WorkoutInfo, garmin_creator: str) -> ET.ElementTree:
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
            ET.SubElement(garmin_trkpt, f"{{{APPLE_GPX_NS}}}time").text = normalize_utc_text(apple_time.text)

    return ET.ElementTree(garmin_root)


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


def indent_xml(tree: ET.ElementTree) -> None:
    try:
        ET.indent(tree, space="  ")
    except AttributeError:
        pass


if __name__ == "__main__":
    main()
