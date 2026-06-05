from __future__ import annotations

import asyncio
import itertools
import os
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, cast

from ._async_process import AsyncStdioRPCProcess
from ._binary import get_binary_path
from ._process import StdioRPCProcess
from ._spreadsheet_base import _AsyncSpreadsheetSessionBase, _SpreadsheetSessionBase, drop_none
from .exceptions import WitanProcessError
from .generated_types import (
    DataValidationInfo,
    DataValidationResult,
    DataValidationSpec,
)


class _WorkbookBase:
    def __init__(
        self,
        path: str | os.PathLike[str],
        *,
        create: bool = False,
        locale: str | None = None,
        hint: str | None = None,
        stateless: bool | None = None,
        api_key: str | None = None,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
    ) -> None:
        self.path = Path(path)
        self.create = create
        self.locale = locale
        self.hint = hint
        self.stateless = stateless
        self.api_key = api_key
        self.api_url = api_url
        self.binary = Path(binary) if binary is not None else None
        self.env = dict(env or {})
        self.request_timeout = request_timeout
        self._ids = itertools.count(1)

    def _next_id(self) -> str:
        return str(next(self._ids))

    def _argv(self) -> list[str]:
        binary = str(self.binary) if self.binary is not None else get_binary_path()
        argv = [binary]
        if self.api_key is not None:
            argv.extend(["--api-key", self.api_key])
        if self.api_url is not None:
            argv.extend(["--api-url", self.api_url])
        if self.stateless is True:
            argv.append("--stateless")
        argv.extend(["xlsx", "rpc", str(self.path)])
        if self.create:
            argv.append("--create")
        if self.hint is not None:
            argv.extend(["--hint", self.hint])
        if self.locale is not None:
            argv.extend(["--locale", self.locale])
        return argv


class Workbook(_WorkbookBase, _SpreadsheetSessionBase):
    """Synchronous Witan workbook session backed by `witan xlsx rpc`."""

    def __init__(
        self,
        path: str | os.PathLike[str],
        *,
        create: bool = False,
        locale: str | None = None,
        hint: str | None = None,
        stateless: bool | None = None,
        api_key: str | None = None,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
    ) -> None:
        _WorkbookBase.__init__(
            self,
            path,
            create=create,
            locale=locale,
            hint=hint,
            stateless=stateless,
            api_key=api_key,
            api_url=api_url,
            binary=binary,
            env=env,
            request_timeout=request_timeout,
        )
        _SpreadsheetSessionBase.__init__(self)
        self._process = StdioRPCProcess(self._argv(), env=self.env, request_timeout=request_timeout)

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        self.close()

    def _request(self, method: str, op: str, args: Mapping[str, Any] | None = None) -> Any:
        return self._process.request(
            request_id=self._next_id(),
            op=op,
            args=args or {},
            method=method,
        )

    def close(self) -> None:
        self._process.close()

    def save(self) -> bool:
        return cast(bool, self._request("save", "save", {}))

    def get_data_validations(self, *, sheet: str | None = None, address: str | None = None) -> list[DataValidationInfo]:
        result = cast(Mapping[str, Any], self._request("get_data_validations", "getDataValidations", drop_none({"sheet": sheet, "address": address})))
        return cast(list[DataValidationInfo], result.get("rules", []))

    def validate_cells(self, address: str, *, max_cells_to_scan: int | None = None, max_invalid_cells: int | None = None, treat_unsupported_as_invalid: bool | None = None) -> DataValidationResult:
        return cast(DataValidationResult, self._request("validate_cells", "validateCells", drop_none({"address": address, "maxCellsToScan": max_cells_to_scan, "maxInvalidCells": max_invalid_cells, "treatUnsupportedAsInvalid": treat_unsupported_as_invalid})))

    def set_data_validations(self, sheet_name: str, rules: Sequence[DataValidationSpec], *, clear: bool | None = None) -> None:
        self._request("set_data_validations", "setDataValidations", drop_none({"sheet": sheet_name, "rules": list(rules), "clear": clear}))

    def remove_data_validations(self, sheet_name: str, *, indices: Sequence[int] | None = None, address: str | None = None) -> None:
        self._request("remove_data_validations", "removeDataValidations", drop_none({"sheet": sheet_name, "indices": list(indices) if indices is not None else None, "address": address}))


class AsyncWorkbook(_WorkbookBase, _AsyncSpreadsheetSessionBase):
    """Asynchronous Witan workbook session backed by `witan xlsx rpc`."""

    def __init__(
        self,
        path: str | os.PathLike[str],
        *,
        create: bool = False,
        locale: str | None = None,
        hint: str | None = None,
        stateless: bool | None = None,
        api_key: str | None = None,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
    ) -> None:
        _WorkbookBase.__init__(
            self,
            path,
            create=create,
            locale=locale,
            hint=hint,
            stateless=stateless,
            api_key=api_key,
            api_url=api_url,
            binary=binary,
            env=env,
            request_timeout=request_timeout,
        )
        _AsyncSpreadsheetSessionBase.__init__(self)
        self._process: AsyncStdioRPCProcess | None = None
        self._start_lock = asyncio.Lock()
        self._closed = False

    async def __aenter__(self) -> "AsyncWorkbook":
        await self._ensure_process()
        return self

    async def __aexit__(self, exc_type: object, exc: object, tb: object) -> None:
        await self.close()

    async def _ensure_process(self) -> AsyncStdioRPCProcess:
        if self._closed:
            raise WitanProcessError("witan subprocess is closed")
        process = self._process
        if process is not None:
            if process.closed:
                raise WitanProcessError("witan subprocess is closed")
            return process

        async with self._start_lock:
            if self._closed:
                raise WitanProcessError("witan subprocess is closed")
            process = self._process
            if process is not None:
                if process.closed:
                    raise WitanProcessError("witan subprocess is closed")
                return process
            self._process = await AsyncStdioRPCProcess.create(self._argv(), env=self.env, request_timeout=self.request_timeout)
            return self._process

    async def _request(self, method: str, op: str, args: Mapping[str, Any] | None = None) -> Any:
        process = await self._ensure_process()
        return await process.request(request_id=self._next_id(), op=op, args=args or {}, method=method)

    async def close(self) -> None:
        async with self._start_lock:
            if self._closed:
                return
            self._closed = True
            if self._process is not None:
                await self._process.close()

    async def save(self) -> bool:
        return cast(bool, await self._request("save", "save", {}))

    async def get_data_validations(self, *, sheet: str | None = None, address: str | None = None) -> list[DataValidationInfo]:
        result = cast(Mapping[str, Any], await self._request("get_data_validations", "getDataValidations", drop_none({"sheet": sheet, "address": address})))
        return cast(list[DataValidationInfo], result.get("rules", []))

    async def validate_cells(self, address: str, *, max_cells_to_scan: int | None = None, max_invalid_cells: int | None = None, treat_unsupported_as_invalid: bool | None = None) -> DataValidationResult:
        return cast(DataValidationResult, await self._request("validate_cells", "validateCells", drop_none({"address": address, "maxCellsToScan": max_cells_to_scan, "maxInvalidCells": max_invalid_cells, "treatUnsupportedAsInvalid": treat_unsupported_as_invalid})))

    async def set_data_validations(self, sheet_name: str, rules: Sequence[DataValidationSpec], *, clear: bool | None = None) -> None:
        await self._request("set_data_validations", "setDataValidations", drop_none({"sheet": sheet_name, "rules": list(rules), "clear": clear}))

    async def remove_data_validations(self, sheet_name: str, *, indices: Sequence[int] | None = None, address: str | None = None) -> None:
        await self._request("remove_data_validations", "removeDataValidations", drop_none({"sheet": sheet_name, "indices": list(indices) if indices is not None else None, "address": address}))
