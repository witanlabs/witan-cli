from __future__ import annotations

import asyncio
import json
import stat
from pathlib import Path

import pytest

from witan import AsyncWorkbook, Regex, WitanProcessError, WitanRPCError, WitanTimeoutError


def fake_binary() -> Path:
    path = Path(__file__).with_name("fake_witan_rpc.py")
    path.chmod(path.stat().st_mode | stat.S_IXUSR)
    return path


def fake_env(tmp_path: Path, *, mode: str = "ok") -> tuple[dict[str, str], Path, Path, Path]:
    argv_file = tmp_path / f"argv-{mode}.jsonl"
    requests_file = tmp_path / f"requests-{mode}.jsonl"
    save_file = tmp_path / f"saved-{mode}.txt"
    env = {
        "WITAN_FAKE_ARGV_FILE": str(argv_file),
        "WITAN_FAKE_REQUESTS_FILE": str(requests_file),
        "WITAN_FAKE_SAVE_FILE": str(save_file),
        "WITAN_FAKE_MODE": mode,
    }
    return env, argv_file, requests_file, save_file


def json_lines(path: Path) -> list[dict[str, object]]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line]


def test_async_workbook_invokes_witan_xlsx_rpc(tmp_path: Path) -> None:
    async def run() -> None:
        env, argv_file, requests_file, save_file = fake_env(tmp_path)
        workbook_path = tmp_path / "async-created.xlsx"
        async with AsyncWorkbook(
            workbook_path,
            create=True,
            stateless=True,
            binary=fake_binary(),
            env=env,
        ) as wb:
            assert await wb.read_range_tsv("Sheet1!A1:B2") == "A\tB\n1\t2"
            assert (await wb.read_cell("Sheet1!A1"))["value"] == 2
            assert await wb.find_cells(Regex("rev", "i")) == []
            assert await wb.reduce_addresses(["Sheet1!A:B"]) == ["Sheet1!A1:B2"]
            assert (await wb.scenarios([{"address": "Sheet1!A1", "values": [1]}], ["Sheet1!B1"]))["sweepCount"] == 1
            assert await wb.save() is True

        assert save_file.read_text(encoding="utf-8") == "saved\n"
        argv = json.loads(argv_file.read_text(encoding="utf-8"))
        assert argv[:4] == ["--stateless", "xlsx", "rpc", str(workbook_path)]
        assert "--create" in argv
        assert [request["op"] for request in json_lines(requests_file)] == [
            "readRangeTsv",
            "readRange",
            "findCells",
            "reduceAddresses",
            "sweepInputs",
            "save",
        ]

    asyncio.run(run())


def test_async_workbook_raises_rpc_error(tmp_path: Path) -> None:
    async def run() -> None:
        env, _, _, _ = fake_env(tmp_path, mode="rpc-error")
        wb = AsyncWorkbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env)
        try:
            with pytest.raises(WitanRPCError) as raised:
                await wb.list_sheets()
        finally:
            await wb.close()
        assert raised.value.code == "BOOM"

    asyncio.run(run())


def test_async_workbook_timeout_terminates_session(tmp_path: Path) -> None:
    async def run() -> None:
        env, _, _, _ = fake_env(tmp_path, mode="hang")
        wb = AsyncWorkbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env, request_timeout=0.05)
        with pytest.raises(WitanTimeoutError):
            await wb.list_sheets()
        with pytest.raises(WitanProcessError):
            await wb.list_sheets()

    asyncio.run(run())


def test_async_workbook_close_prevents_reuse(tmp_path: Path) -> None:
    async def run() -> None:
        env, _, requests_file, _ = fake_env(tmp_path)
        wb = AsyncWorkbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env)
        assert await wb.list_sheets()
        await wb.close()

        with pytest.raises(WitanProcessError):
            await wb.list_sheets()

        assert [request["op"] for request in json_lines(requests_file)] == ["listSheets"]

    asyncio.run(run())


def test_async_workbook_close_before_start_prevents_use(tmp_path: Path) -> None:
    async def run() -> None:
        env, _, requests_file, _ = fake_env(tmp_path)
        wb = AsyncWorkbook(tmp_path / "book.xlsx", binary=fake_binary(), env=env)
        await wb.close()

        with pytest.raises(WitanProcessError):
            await wb.list_sheets()

        assert not requests_file.exists()

    asyncio.run(run())


def test_async_workbook_serializes_first_startup(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    async def run() -> None:
        created = 0

        class FakeProcess:
            closed = False

            async def request(self, **_kwargs: object) -> object:
                return {"sheets": []}

            async def close(self) -> None:
                self.closed = True

        async def create_process(*_args: object, **_kwargs: object) -> FakeProcess:
            nonlocal created
            created += 1
            await asyncio.sleep(0)
            return FakeProcess()

        monkeypatch.setattr("witan.workbook.AsyncStdioRPCProcess.create", create_process)

        wb = AsyncWorkbook(tmp_path / "book.xlsx", binary=fake_binary())
        try:
            await asyncio.gather(wb.list_sheets(), wb.list_sheets())
        finally:
            await wb.close()

        assert created == 1

    asyncio.run(run())
