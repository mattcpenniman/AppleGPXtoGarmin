from __future__ import annotations

from getpass import getpass
import os
from pathlib import Path
import re
from typing import Any


def create_client(email: str | None, token_store: Path) -> Any:
    try:
        from garminconnect import (
            Garmin,
            GarminConnectAuthenticationError,
            GarminConnectConnectionError,
        )
    except ImportError as exc:
        raise RuntimeError(
            "Garmin export requires the project dependencies; run "
            "'python -m pip install -e .' first"
        ) from exc

    token_path = str(token_store.expanduser())
    try:
        client = Garmin()
        client.login(token_path)
        return client
    except (
        FileNotFoundError,
        GarminConnectAuthenticationError,
        GarminConnectConnectionError,
    ):
        pass

    login_email = email or os.getenv("GARMIN_EMAIL") or os.getenv("EMAIL")
    if not login_email:
        login_email = input("Garmin email: ").strip()
    password = os.getenv("GARMIN_PASSWORD") or os.getenv("PASSWORD") or getpass(
        "Garmin password: "
    )
    if not login_email or not password:
        raise ValueError("Garmin email and password are required")

    client = Garmin(
        email=login_email,
        password=password,
        prompt_mfa=lambda: input("Garmin MFA code: ").strip(),
    )
    client.login(token_path)
    return client


def select_activities(
    client: Any,
    activity_ids: list[str],
    start_date: str | None,
    end_date: str | None,
    activity_type: str | None,
    limit: int,
) -> list[dict[str, Any]]:
    if activity_ids:
        return [{"activityId": activity_id} for activity_id in activity_ids]

    if start_date:
        activities = client.get_activities_by_date(
            start_date,
            end_date,
            activity_type,
            "desc",
        )
        return activities[:limit]

    activities = client.get_activities(0, limit, activity_type)
    if isinstance(activities, dict):
        activities = activities.get("activityList", [])
    return list(activities)


def export_tcx(
    client: Any,
    activities: list[dict[str, Any]],
    output_dir: Path,
    overwrite: bool,
) -> dict[str, int]:
    try:
        from garminconnect import Garmin
    except ImportError as exc:
        raise RuntimeError("Garmin export requires the garminconnect package") from exc

    output_dir.mkdir(parents=True, exist_ok=True)
    summary = {"written": 0, "skipped_existing": 0}

    for activity in activities:
        activity_id = str(activity.get("activityId", "")).strip()
        if not activity_id:
            raise ValueError("Garmin returned an activity without an activityId")

        name = sanitize_filename(str(activity.get("activityName", "")))
        suffix = f"_{name}" if name else ""
        output_path = output_dir / f"garmin_{activity_id}{suffix}.tcx"
        if output_path.exists() and not overwrite:
            summary["skipped_existing"] += 1
            continue

        content = client.download_activity(
            activity_id,
            dl_fmt=Garmin.ActivityDownloadFormat.TCX,
        )
        output_path.write_bytes(content)
        summary["written"] += 1

    return summary


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned.strip("._-")[:80]
