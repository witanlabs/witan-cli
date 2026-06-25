from __future__ import annotations

import asyncio
import json
import os
import subprocess
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from ._async_process import AsyncStdioRPCProcess
from ._binary import get_binary_path
from ._process import StdioRPCProcess
from ._spreadsheet_base import _AsyncSpreadsheetSessionBase, _SpreadsheetSessionBase
from .exceptions import WitanProcessError
_DEFAULT_REQUEST_TIMEOUT = 180.0
# Explicit "create a new sheet" sentinels. An empty ref is NOT included here:
# it must not silently infer create mode (e.g. an unset env var / config value).
_CREATE_REFS = frozenset({"new", "gs://new"})

# The authorize/status CLI commands cap their internal poll at 5 minutes; allow
# a little headroom for the subprocess round trip.
_AUTHORIZE_WAIT_TIMEOUT = 360.0


def _cli_argv(
    extra: list[str],
    *,
    binary: str | os.PathLike[str] | None,
    api_url: str | None,
) -> list[str]:
    argv = [str(binary) if binary is not None else get_binary_path()]
    if api_url is not None:
        argv.extend(["--api-url", api_url])
    argv.extend(extra)
    return argv


def _run_cli_json(
    extra: list[str],
    *,
    binary: str | os.PathLike[str] | None = None,
    api_url: str | None = None,
    env: Mapping[str, str] | None = None,
    timeout: float = 60.0,
) -> dict[str, Any]:
    """Run a one-shot `witan` subcommand with --json and parse its stdout."""
    argv = _cli_argv(extra, binary=binary, api_url=api_url)
    merged_env = os.environ.copy()
    if env:
        merged_env.update(env)
    try:
        proc = subprocess.run(argv, capture_output=True, env=merged_env, timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        raise WitanProcessError(f"witan {' '.join(extra)} timed out after {timeout:g}s") from exc

    stdout = proc.stdout.decode("utf-8", errors="replace").strip()
    if proc.returncode != 0:
        stderr = proc.stderr.decode("utf-8", errors="replace").strip()
        raise WitanProcessError(
            f"witan {' '.join(extra)} failed (exit {proc.returncode})",
            stderr_tail=tuple(stderr.splitlines()[-10:]),
        )
    if not stdout:
        return {}
    try:
        parsed = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise WitanProcessError(f"invalid JSON from witan {' '.join(extra)}: {stdout!r}") from exc
    if not isinstance(parsed, dict):
        raise WitanProcessError(f"expected JSON object from witan {' '.join(extra)}: {stdout!r}")
    return parsed


def _authorize_url(ref: str, **kwargs: Any) -> dict[str, Any]:
    return _run_cli_json(["gsheets", "authorize", ref, "--json"], **kwargs)


def _is_authorized(ref: str, **kwargs: Any) -> bool:
    data = _run_cli_json(["gsheets", "status", ref, "--json"], **kwargs)
    return bool(data.get("authorized"))


def _wait_until_authorized(ref: str, *, timeout: float = _AUTHORIZE_WAIT_TIMEOUT, **kwargs: Any) -> bool:
    data = _run_cli_json(["gsheets", "status", ref, "--wait", "--json"], timeout=timeout, **kwargs)
    return bool(data.get("authorized"))


def _is_create_ref(ref: str) -> bool:
    return ref in _CREATE_REFS


def _reject_google_sheet_options(
    *,
    api_key: str | None,
    stateless: bool | None,
    hint: str | None,
) -> None:
    if api_key is not None:
        msg = "Google Sheets requires user authentication. Do not pass api_key; run 'witan auth login' instead."
        raise ValueError(msg)
    if stateless is not None:
        msg = "GoogleSheet does not support stateless mode."
        raise ValueError(msg)
    if hint is not None:
        msg = "GoogleSheet does not support hint."
        raise ValueError(msg)


def _resolve_create(ref: str, create: bool) -> bool:
    if create:
        return True
    return _is_create_ref(ref)


def _validate_ref_create(ref: str, create: bool) -> None:
    if create:
        # An empty ref is allowed only in create mode (GoogleSheet.create() omits
        # it); a non-empty ref must be an explicit "new" sentinel.
        if ref and ref not in _CREATE_REFS:
            msg = "create requires ref 'new' or gs://new, or omit ref (use GoogleSheet.create())"
            raise ValueError(msg)
        return
    if not ref:
        msg = "ref is required when opening an existing spreadsheet"
        raise ValueError(msg)


class _GoogleSheetSessionConfig:
    def __init__(
        self,
        ref: str = "",
        *,
        create: bool = False,
        title: str | None = None,
        locale: str | None = None,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
        api_key: str | None = None,
        stateless: bool | None = None,
        hint: str | None = None,
    ) -> None:
        _reject_google_sheet_options(api_key=api_key, stateless=stateless, hint=hint)
        self._rpc_create = _resolve_create(ref, create)
        _validate_ref_create(ref, self._rpc_create)
        self.ref = ref
        self.title = title
        self.locale = locale
        self.api_url = api_url
        self.binary = Path(binary) if binary is not None else None
        self.env = dict(env or {})
        self.request_timeout = _DEFAULT_REQUEST_TIMEOUT if request_timeout is None else request_timeout

    def _argv(self) -> list[str]:
        binary = str(self.binary) if self.binary is not None else get_binary_path()
        argv = [binary]
        if self.api_url is not None:
            argv.extend(["--api-url", self.api_url])
        argv.extend(["gsheets", "rpc"])
        if self._rpc_create:
            argv.append("--create")
            if self.title is not None:
                argv.extend(["--title", self.title])
            if self.ref in ("new", "gs://new"):
                argv.append(self.ref)
        else:
            argv.append(self.ref)
        if self.locale is not None:
            argv.extend(["--locale", self.locale])
        return argv


class GoogleSheet(_GoogleSheetSessionConfig, _SpreadsheetSessionBase):
    """Synchronous Google Sheets session backed by `witan gsheets rpc`.

    Requires prior CLI setup: ``witan auth login`` and ``witan gsheets connect``.
    Changes persist immediately; there is no ``save()`` method.

    Open an existing spreadsheet by URL or ``gs://`` reference. Create a new one with
    :meth:`create` or ``GoogleSheet("new", title=...)``.
    """

    def __init__(
        self,
        ref: str = "",
        *,
        create: bool = False,
        title: str | None = None,
        locale: str | None = None,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
        api_key: str | None = None,
        stateless: bool | None = None,
        hint: str | None = None,
    ) -> None:
        _GoogleSheetSessionConfig.__init__(
            self,
            ref,
            create=create,
            title=title,
            locale=locale,
            api_url=api_url,
            binary=binary,
            env=env,
            request_timeout=request_timeout,
            api_key=api_key,
            stateless=stateless,
            hint=hint,
        )
        _SpreadsheetSessionBase.__init__(self)
        self._process = StdioRPCProcess(self._argv(), env=self.env, request_timeout=self.request_timeout)

    @classmethod
    def create(
        cls,
        title: str | None = None,
        *,
        locale: str | None = None,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
        api_key: str | None = None,
    ) -> "GoogleSheet":
        """Create a new Google spreadsheet and open an RPC session."""
        return cls(
            create=True,
            title=title,
            locale=locale,
            api_url=api_url,
            binary=binary,
            env=env,
            request_timeout=request_timeout,
            api_key=api_key,
        )

    @staticmethod
    def authorize_url(
        ref: str,
        *,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
    ) -> dict[str, Any]:
        """Begin per-file authorization for *ref*.

        Returns ``{"authorized": True}`` if the sheet is already authorized,
        otherwise ``{"authorized": False, "picker_url": ..., "expires_in_seconds": ...}``.
        Hand ``picker_url`` to a human to open Google's file picker, then call
        :meth:`wait_until_authorized`.
        """
        return _authorize_url(ref, api_url=api_url, binary=binary, env=env)

    @staticmethod
    def is_authorized(
        ref: str,
        *,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
    ) -> bool:
        """Return whether *ref* is authorized for the app."""
        return _is_authorized(ref, api_url=api_url, binary=binary, env=env)

    @staticmethod
    def wait_until_authorized(
        ref: str,
        *,
        timeout: float = _AUTHORIZE_WAIT_TIMEOUT,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
    ) -> bool:
        """Block until *ref* is authorized. Returns True; raises on timeout."""
        return _wait_until_authorized(ref, timeout=timeout, api_url=api_url, binary=binary, env=env)

    def __enter__(self) -> "GoogleSheet":
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

    @property
    def is_create(self) -> bool:
        """True when this session was opened in create mode."""
        return self._rpc_create


class AsyncGoogleSheet(_GoogleSheetSessionConfig, _AsyncSpreadsheetSessionBase):
    """Asynchronous Google Sheets session backed by `witan gsheets rpc`.

    Requires prior CLI setup: ``witan auth login`` and ``witan gsheets connect``.
    Changes persist immediately; there is no ``save()`` method.
    """

    def __init__(
        self,
        ref: str = "",
        *,
        create: bool = False,
        title: str | None = None,
        locale: str | None = None,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
        api_key: str | None = None,
        stateless: bool | None = None,
        hint: str | None = None,
    ) -> None:
        _GoogleSheetSessionConfig.__init__(
            self,
            ref,
            create=create,
            title=title,
            locale=locale,
            api_url=api_url,
            binary=binary,
            env=env,
            request_timeout=request_timeout,
            api_key=api_key,
            stateless=stateless,
            hint=hint,
        )
        _AsyncSpreadsheetSessionBase.__init__(self)
        self._process: AsyncStdioRPCProcess | None = None
        self._start_lock = asyncio.Lock()
        self._closed = False

    @classmethod
    def create(
        cls,
        title: str | None = None,
        *,
        locale: str | None = None,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
        api_key: str | None = None,
    ) -> "AsyncGoogleSheet":
        """Create a new Google spreadsheet and open an async RPC session."""
        return cls(
            create=True,
            title=title,
            locale=locale,
            api_url=api_url,
            binary=binary,
            env=env,
            request_timeout=request_timeout,
            api_key=api_key,
        )

    @staticmethod
    async def authorize_url(
        ref: str,
        *,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
    ) -> dict[str, Any]:
        """Async variant of :meth:`GoogleSheet.authorize_url`."""
        return await asyncio.to_thread(_authorize_url, ref, api_url=api_url, binary=binary, env=env)

    @staticmethod
    async def is_authorized(
        ref: str,
        *,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
    ) -> bool:
        """Async variant of :meth:`GoogleSheet.is_authorized`."""
        return await asyncio.to_thread(_is_authorized, ref, api_url=api_url, binary=binary, env=env)

    @staticmethod
    async def wait_until_authorized(
        ref: str,
        *,
        timeout: float = _AUTHORIZE_WAIT_TIMEOUT,
        api_url: str | None = None,
        binary: str | os.PathLike[str] | None = None,
        env: Mapping[str, str] | None = None,
    ) -> bool:
        """Async variant of :meth:`GoogleSheet.wait_until_authorized`."""
        return await asyncio.to_thread(
            _wait_until_authorized, ref, timeout=timeout, api_url=api_url, binary=binary, env=env
        )

    async def __aenter__(self) -> "AsyncGoogleSheet":
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
            self._process = await AsyncStdioRPCProcess.create(
                self._argv(),
                env=self.env,
                request_timeout=self.request_timeout,
            )
            return self._process

    async def _request(self, method: str, op: str, args: Mapping[str, Any] | None = None) -> Any:
        process = await self._ensure_process()
        return await process.request(
            request_id=self._next_id(),
            op=op,
            args=args or {},
            method=method,
        )

    async def close(self) -> None:
        async with self._start_lock:
            if self._closed:
                return
            self._closed = True
            if self._process is not None:
                await self._process.close()

    @property
    def is_create(self) -> bool:
        """True when this session was opened in create mode."""
        return self._rpc_create
