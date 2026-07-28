from __future__ import annotations

import argparse
from dataclasses import replace
import os
from pathlib import Path
import sys

from apple_to_garmin_gpx import (
    convert_routes,
    load_config,
    load_running_workout_routes,
)

from .garmin import create_client, export_tcx, select_activities


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="apple-garmin",
        description="Convert Apple Health activities and work with Garmin Connect.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    convert = subparsers.add_parser(
        "convert", help="Convert Apple Health running workouts to GPX or TCX"
    )
    convert.add_argument("--env-file", type=Path, default=Path(".env"))
    convert.add_argument("--apple-export-dir", type=Path)
    convert.add_argument("--export-xml", type=Path)
    convert.add_argument("--workout-routes-dir", type=Path)
    convert.add_argument("--output-dir", type=Path)
    convert.add_argument("--format", choices=("gpx", "tcx"), dest="mode")
    convert.add_argument("--limit", type=positive_int)
    convert.add_argument("--fuzzy-match-seconds", type=non_negative_int)
    convert.add_argument(
        "--debug-xlsx", action=argparse.BooleanOptionalAction, default=None
    )
    convert.add_argument(
        "--overwrite", action=argparse.BooleanOptionalAction, default=None
    )
    convert.set_defaults(handler=run_convert)

    garmin = subparsers.add_parser("garmin", help="Use the unofficial Garmin API")
    garmin_subparsers = garmin.add_subparsers(dest="garmin_command", required=True)
    download = garmin_subparsers.add_parser(
        "export-tcx", help="Download Garmin Connect activities as TCX"
    )
    source = download.add_mutually_exclusive_group()
    source.add_argument(
        "--activity-id",
        action="append",
        default=[],
        help="Activity ID to export; repeat for multiple activities",
    )
    source.add_argument("--start-date", help="First activity date (YYYY-MM-DD)")
    download.add_argument("--end-date", help="Last activity date (YYYY-MM-DD)")
    download.add_argument("--activity-type", help="Garmin activity type filter")
    download.add_argument(
        "--limit", type=positive_int, default=1, help="Maximum activities (default: 1)"
    )
    download.add_argument("--output-dir", type=Path, default=Path("garmin_exports"))
    download.add_argument("--email", help="Garmin email; password is prompted securely")
    download.add_argument(
        "--token-store",
        type=Path,
        default=Path(os.getenv("GARMINTOKENS", "~/.garminconnect")).expanduser(),
        help="OAuth token directory (default: ~/.garminconnect)",
    )
    download.add_argument("--overwrite", action="store_true")
    download.set_defaults(handler=run_garmin_export)

    return parser


def run_convert(args: argparse.Namespace) -> int:
    env_path = args.env_file.expanduser().resolve()
    config = load_config(env_path.parent, env_path)
    overrides = {
        "apple_export_dir": resolve_cli_path(args.apple_export_dir),
        "export_xml_path": resolve_cli_path(args.export_xml),
        "workout_routes_dir": resolve_cli_path(args.workout_routes_dir),
        "output_dir": resolve_cli_path(args.output_dir),
        "mode": args.mode,
        "limit": args.limit,
        "fuzzy_match_seconds": args.fuzzy_match_seconds,
        "debug_xlsx": args.debug_xlsx,
        "overwrite_existing": args.overwrite,
    }
    config = replace(
        config,
        **{key: value for key, value in overrides.items() if value is not None},
    )
    route_map = load_running_workout_routes(config.export_xml_path)
    summary = convert_routes(config, route_map)
    print(f"Running workouts found: {len(route_map)}")
    print(f"{config.mode.upper()} files written: {summary['written']}")
    print(f"Existing files skipped: {summary['skipped_existing']}")
    print(f"Missing route files: {summary['missing_routes']}")
    print(f"Output folder: {config.output_dir}")
    return 0


def run_garmin_export(args: argparse.Namespace) -> int:
    if args.end_date and not args.start_date:
        raise ValueError("--end-date requires --start-date")

    client = create_client(args.email, args.token_store)
    activities = select_activities(
        client=client,
        activity_ids=args.activity_id,
        start_date=args.start_date,
        end_date=args.end_date,
        activity_type=args.activity_type,
        limit=args.limit,
    )
    summary = export_tcx(
        client=client,
        activities=activities,
        output_dir=args.output_dir.expanduser().resolve(),
        overwrite=args.overwrite,
    )
    print(f"Activities selected: {len(activities)}")
    print(f"TCX files written: {summary['written']}")
    print(f"Existing files skipped: {summary['skipped_existing']}")
    print(f"Output folder: {args.output_dir.expanduser().resolve()}")
    return 0


def resolve_cli_path(value: Path | None) -> Path | None:
    return value.expanduser().resolve() if value is not None else None


def positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def non_negative_int(value: str) -> int:
    parsed = int(value)
    if parsed < 0:
        raise argparse.ArgumentTypeError("must be zero or greater")
    return parsed


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.handler(args)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
