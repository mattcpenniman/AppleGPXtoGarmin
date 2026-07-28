from pathlib import Path
import sys
from types import SimpleNamespace

from apple_garmin.garmin import export_tcx, sanitize_filename, select_activities


class FakeClient:
    def __init__(self) -> None:
        self.downloaded: list[tuple[str, object]] = []

    def get_activities(self, start: int, limit: int, activity_type: str | None):
        assert (start, limit, activity_type) == (0, 1, "running")
        return [{"activityId": 42, "activityName": "Morning Run"}]

    def get_activities_by_date(
        self,
        start: str,
        end: str | None,
        activity_type: str | None,
        order: str,
    ):
        assert (start, end, activity_type, order) == (
            "2026-01-01",
            "2026-01-31",
            "running",
            "desc",
        )
        return [{"activityId": index} for index in range(3)]

    def download_activity(self, activity_id: str, dl_fmt: object) -> bytes:
        self.downloaded.append((activity_id, dl_fmt))
        return f"tcx-{activity_id}".encode()


def test_selects_latest_or_date_range() -> None:
    client = FakeClient()
    latest = select_activities(client, [], None, None, "running", 1)
    ranged = select_activities(
        client, [], "2026-01-01", "2026-01-31", "running", 2
    )

    assert latest[0]["activityId"] == 42
    assert [item["activityId"] for item in ranged] == [0, 1]


def test_exports_tcx_and_respects_existing_file(tmp_path: Path, monkeypatch) -> None:
    tcx_format = object()
    fake_garmin = SimpleNamespace(ActivityDownloadFormat=SimpleNamespace(TCX=tcx_format))
    monkeypatch.setitem(sys.modules, "garminconnect", SimpleNamespace(Garmin=fake_garmin))
    client = FakeClient()
    activities = [{"activityId": 42, "activityName": "Morning / Run"}]

    first = export_tcx(client, activities, tmp_path, overwrite=False)
    second = export_tcx(client, activities, tmp_path, overwrite=False)

    output = tmp_path / "garmin_42_Morning_Run.tcx"
    assert output.read_bytes() == b"tcx-42"
    assert first == {"written": 1, "skipped_existing": 0}
    assert second == {"written": 0, "skipped_existing": 1}
    assert client.downloaded == [("42", tcx_format)]


def test_sanitize_filename() -> None:
    assert sanitize_filename("  Run: 5 km / easy  ") == "Run_5_km_easy"
