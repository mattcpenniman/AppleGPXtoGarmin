import pytest

from apple_garmin.cli import build_parser, non_negative_int, positive_int


def test_parses_garmin_export_command() -> None:
    args = build_parser().parse_args(
        ["garmin", "export-tcx", "--activity-id", "123", "--overwrite"]
    )

    assert args.activity_id == ["123"]
    assert args.overwrite is True


@pytest.mark.parametrize("value", ["0", "-1"])
def test_positive_int_rejects_non_positive_values(value: str) -> None:
    with pytest.raises(Exception):
        positive_int(value)


def test_non_negative_int_accepts_zero() -> None:
    assert non_negative_int("0") == 0
