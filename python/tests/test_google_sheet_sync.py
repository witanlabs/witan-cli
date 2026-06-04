from __future__ import annotations

import json
import stat
from pathlib import Path

import pytest

from witan import GoogleSheet, WitanProcessError, WitanRPCError, is_google_auth_required
from witan.google_sheet import _DEFAULT_REQUEST_TIMEOUT


def fake_binary() -> Path:
    path = Path(__file__).with_name("fake_witan_rpc.py")
    path.chmod(path.stat().st_mode | stat.S_IXUSR)
    return path


def json_lines(path: Path) -> list[dict[str, object]]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line]


def fake_env(tmp_path: Path, *, mode: str = "ok") -> tuple[dict[str, str], Path, Path]:
    argv_file = tmp_path / "argv.jsonl"
    requests_file = tmp_path / "requests.jsonl"
    env = {
        "WITAN_FAKE_ARGV_FILE": str(argv_file),
        "WITAN_FAKE_REQUESTS_FILE": str(requests_file),
        "WITAN_FAKE_MODE": mode,
    }
    return env, argv_file, requests_file


def test_google_sheet_open_uses_gsheets_rpc(tmp_path: Path) -> None:
    env, argv_file, requests_file = fake_env(tmp_path)

    with GoogleSheet("gs://sheet-123", binary=fake_binary(), env=env) as sheet:
        assert sheet.list_sheets()[0]["sheet"] == "Sheet1"
        assert sheet.read_range("Sheet1!A1:B2")[0][0]["value"] == 2
        assert sheet.set_cells([{"address": "Sheet1!A1", "value": "done"}])["changed"] == ["Sheet1!A1"]
        assert not hasattr(sheet, "save")

    argv = json.loads(argv_file.read_text(encoding="utf-8"))
    assert argv == ["gsheets", "rpc", "gs://sheet-123"]

    requests = json_lines(requests_file)
    assert [request["op"] for request in requests] == ["listSheets", "readRange", "setCells"]


def test_google_sheet_create_uses_single_rpc_with_create_flag(tmp_path: Path) -> None:
    env, argv_file, requests_file = fake_env(tmp_path)

    with GoogleSheet.create("Budget 2026", binary=fake_binary(), env=env) as sheet:
        assert sheet.is_create is True
        assert sheet.read_cell("Sheet1!A1")["value"] == 2

    argv = json.loads(argv_file.read_text(encoding="utf-8"))
    assert argv == ["gsheets", "rpc", "--create", "--title", "Budget 2026"]

    assert json_lines(requests_file)[0]["op"] == "readRange"


def test_google_sheet_create_via_new_ref(tmp_path: Path) -> None:
    env, argv_file, requests_file = fake_env(tmp_path)

    with GoogleSheet("new", title="Q1", binary=fake_binary(), env=env) as sheet:
        assert sheet.is_create is True
        assert sheet.read_range("Sheet1!A1:B2")

    argv = json.loads(argv_file.read_text(encoding="utf-8"))
    assert argv == ["gsheets", "rpc", "--create", "--title", "Q1", "new"]

    assert json_lines(requests_file)


def test_google_sheet_default_timeout(tmp_path: Path) -> None:
    env, _, _ = fake_env(tmp_path)
    sheet = GoogleSheet("gs://sheet-123", binary=fake_binary(), env=env)
    try:
        assert sheet.request_timeout == _DEFAULT_REQUEST_TIMEOUT
        assert sheet._process._timeout == _DEFAULT_REQUEST_TIMEOUT
    finally:
        sheet.close()


def test_google_sheet_rejects_xlsx_only_options() -> None:
    with pytest.raises(ValueError, match="api_key"):
        GoogleSheet("gs://id", api_key="secret")

    with pytest.raises(ValueError, match="stateless"):
        GoogleSheet("gs://id", stateless=True)

    with pytest.raises(ValueError, match="create requires"):
        GoogleSheet("gs://id", create=True)

    with pytest.raises(ValueError, match="hint"):
        GoogleSheet("gs://id", hint="Sheet1!A1")

    with pytest.raises(ValueError, match="api_key"):
        GoogleSheet.create(api_key="secret")


def test_google_sheet_rpc_error_code(tmp_path: Path) -> None:
    env, _, _ = fake_env(tmp_path, mode="rpc-error")
    sheet = GoogleSheet("gs://sheet-123", binary=fake_binary(), env=env)
    try:
        with pytest.raises(WitanRPCError) as raised:
            sheet.list_sheets()
        assert raised.value.code == "BOOM"
        assert is_google_auth_required(raised.value) is False
    finally:
        sheet.close()


def test_is_google_auth_required() -> None:
    err = WitanRPCError(
        "auth",
        method="list_sheets",
        op="listSheets",
        request_id="1",
        code="google_auth_required",
    )
    assert is_google_auth_required(err) is True


def test_google_sheet_create_failure(tmp_path: Path) -> None:
    env, _, _ = fake_env(tmp_path)

    with pytest.raises(WitanProcessError):
        GoogleSheet.create(title="Nope", binary="/bin/false", env=env)
