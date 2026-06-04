from __future__ import annotations

import asyncio
import json
import stat
from pathlib import Path

import pytest

from witan import AsyncGoogleSheet, WitanProcessError, WitanRPCError, is_google_auth_required
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


def test_async_google_sheet_open_uses_gsheets_rpc(tmp_path: Path) -> None:
    async def run() -> None:
        env, argv_file, requests_file = fake_env(tmp_path)

        async with AsyncGoogleSheet("gs://sheet-123", binary=fake_binary(), env=env) as sheet:
            sheets = await sheet.list_sheets()
            assert sheets[0]["sheet"] == "Sheet1"
            data = await sheet.read_range("Sheet1!A1:B2")
            assert data[0][0]["value"] == 2
            result = await sheet.set_cells([{"address": "Sheet1!A1", "value": "done"}])
            assert result["changed"] == ["Sheet1!A1"]
            assert not hasattr(sheet, "save")

        argv = json.loads(argv_file.read_text(encoding="utf-8"))
        assert argv == ["gsheets", "rpc", "gs://sheet-123"]

        requests = json_lines(requests_file)
        assert [request["op"] for request in requests] == ["listSheets", "readRange", "setCells"]

    asyncio.run(run())


def test_async_google_sheet_create_uses_single_rpc_with_create_flag(tmp_path: Path) -> None:
    async def run() -> None:
        env, argv_file, requests_file = fake_env(tmp_path)

        async with AsyncGoogleSheet.create("Budget 2026", binary=fake_binary(), env=env) as sheet:
            assert sheet.is_create is True
            cell = await sheet.read_cell("Sheet1!A1")
            assert cell["value"] == 2

        argv = json.loads(argv_file.read_text(encoding="utf-8"))
        assert argv == ["gsheets", "rpc", "--create", "--title", "Budget 2026"]

        assert json_lines(requests_file)[0]["op"] == "readRange"

    asyncio.run(run())


def test_async_google_sheet_default_timeout(tmp_path: Path) -> None:
    async def run() -> None:
        env, _, _ = fake_env(tmp_path)
        sheet = AsyncGoogleSheet("gs://sheet-123", binary=fake_binary(), env=env)
        try:
            assert sheet.request_timeout == _DEFAULT_REQUEST_TIMEOUT
            process = await sheet._ensure_process()
            assert process._timeout == _DEFAULT_REQUEST_TIMEOUT
        finally:
            await sheet.close()

    asyncio.run(run())


def test_async_google_sheet_rejects_xlsx_only_options() -> None:
    with pytest.raises(ValueError, match="api_key"):
        AsyncGoogleSheet("gs://id", api_key="secret")

    with pytest.raises(ValueError, match="stateless"):
        AsyncGoogleSheet("gs://id", stateless=True)

    with pytest.raises(ValueError, match="hint"):
        AsyncGoogleSheet("gs://id", hint="Sheet1!A1")


def test_async_google_sheet_rpc_error_code(tmp_path: Path) -> None:
    async def run() -> None:
        env, _, _ = fake_env(tmp_path, mode="rpc-error")
        sheet = AsyncGoogleSheet("gs://sheet-123", binary=fake_binary(), env=env)
        try:
            with pytest.raises(WitanRPCError) as raised:
                await sheet.list_sheets()
            assert raised.value.code == "BOOM"
            assert is_google_auth_required(raised.value) is False
        finally:
            await sheet.close()

    asyncio.run(run())


def test_async_google_sheet_create_failure(tmp_path: Path) -> None:
    async def run() -> None:
        env, _, _ = fake_env(tmp_path)
        sheet = AsyncGoogleSheet.create(title="Nope", binary="/bin/false", env=env)
        try:
            with pytest.raises(WitanProcessError):
                await sheet.list_sheets()
        finally:
            await sheet.close()

    asyncio.run(run())
