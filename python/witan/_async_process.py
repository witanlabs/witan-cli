from __future__ import annotations

import asyncio
import contextlib
import json
import os
from collections import deque
from collections.abc import Mapping
from typing import Any

from .exceptions import WitanProcessError, WitanRPCError, WitanTimeoutError


class AsyncStdioRPCProcess:
    def __init__(
        self,
        argv: list[str],
        *,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
    ) -> None:
        merged_env = os.environ.copy()
        if env:
            merged_env.update(env)
        self._argv = argv
        self._env = merged_env
        self._timeout = 90.0 if request_timeout is None else request_timeout
        self._stderr: deque[str] = deque(maxlen=50)
        self._lock = asyncio.Lock()
        self._proc: asyncio.subprocess.Process | None = None
        self._stderr_task: asyncio.Task[None] | None = None
        self._closed = False

    @classmethod
    async def create(
        cls,
        argv: list[str],
        *,
        env: Mapping[str, str] | None = None,
        request_timeout: float | None = None,
    ) -> "AsyncStdioRPCProcess":
        proc = cls(argv, env=env, request_timeout=request_timeout)
        await proc.start()
        return proc

    async def start(self) -> None:
        try:
            self._proc = await asyncio.create_subprocess_exec(
                *self._argv,
                stdin=asyncio.subprocess.PIPE,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                env=self._env,
            )
        except OSError as exc:
            raise WitanProcessError(f"starting witan subprocess: {exc}") from exc
        if self._proc.stdin is None or self._proc.stdout is None or self._proc.stderr is None:
            await self.terminate()
            raise WitanProcessError("starting witan subprocess: stdio pipes unavailable")
        self._stderr_task = asyncio.create_task(self._drain_stderr())

    async def _drain_stderr(self) -> None:
        assert self._proc is not None and self._proc.stderr is not None
        while True:
            line = await self._proc.stderr.readline()
            if not line:
                return
            text = line.decode(errors="replace").rstrip("\r\n")
            if text:
                self._stderr.append(text)

    def _stderr_tail(self) -> tuple[str, ...]:
        return tuple(self._stderr)

    async def request(self, *, request_id: str, op: str, args: Mapping[str, Any], method: str) -> Any:
        async with self._lock:
            if self._closed:
                raise WitanProcessError("witan subprocess is closed", stderr_tail=self._stderr_tail())
            assert self._proc is not None and self._proc.stdin is not None and self._proc.stdout is not None
            payload = json.dumps(
                {"id": request_id, "op": op, "args": dict(args)},
                separators=(",", ":"),
            )
            try:
                self._proc.stdin.write((payload + "\n").encode())
                await self._proc.stdin.drain()
                line = await asyncio.wait_for(self._proc.stdout.readline(), timeout=self._timeout)
            except asyncio.TimeoutError as exc:
                await self.terminate()
                raise WitanTimeoutError(
                    f"RPC timeout: {method} ({op}) did not respond within {self._timeout:g}s",
                    stderr_tail=self._stderr_tail(),
                ) from exc
            except asyncio.CancelledError:
                await self.terminate()
                raise
            except OSError as exc:
                await self.terminate()
                raise WitanProcessError(f"RPC subprocess I/O failed: {exc}", stderr_tail=self._stderr_tail()) from exc

            if not line:
                code = self._proc.returncode
                raise WitanProcessError(
                    f"witan subprocess exited before responding (exit={code})",
                    stderr_tail=self._stderr_tail(),
                )
            text = line.decode(errors="replace").rstrip("\r\n")
            try:
                response = json.loads(text)
            except json.JSONDecodeError as exc:
                await self.terminate()
                raise WitanProcessError(f"invalid JSON RPC response: {text!r}", stderr_tail=self._stderr_tail()) from exc

            if not isinstance(response, dict):
                await self.terminate()
                raise WitanProcessError(f"invalid RPC response object: {response!r}", stderr_tail=self._stderr_tail())

            response_id = response.get("id")
            if response_id not in (None, "", request_id):
                await self.terminate()
                raise WitanProcessError(
                    f"RPC response id mismatch: expected {request_id!r}, got {response_id!r}",
                    stderr_tail=self._stderr_tail(),
                )

            if response.get("ok") is False:
                raise WitanRPCError(
                    str(response.get("message") or "RPC request failed"),
                    method=method,
                    op=op,
                    request_id=request_id,
                    code=str(response.get("code")) if response.get("code") is not None else None,
                    response=response,
                )
            if response.get("ok") is not True:
                await self.terminate()
                raise WitanProcessError(f"invalid RPC ok field: {response!r}", stderr_tail=self._stderr_tail())
            return response.get("result")

    async def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        if self._proc is None:
            return
        if self._proc.stdin is not None:
            with contextlib.suppress(OSError):
                self._proc.stdin.close()
                await self._proc.stdin.wait_closed()
        try:
            await asyncio.wait_for(self._proc.wait(), timeout=5)
        except asyncio.TimeoutError:
            await self.terminate()
        if self._stderr_task is not None:
            await asyncio.gather(self._stderr_task, return_exceptions=True)

    async def terminate(self) -> None:
        self._closed = True
        if self._proc is None:
            return
        if self._proc.returncode is None:
            self._proc.terminate()
            try:
                await asyncio.wait_for(self._proc.wait(), timeout=2)
            except asyncio.TimeoutError:
                self._proc.kill()
                await self._proc.wait()
        if self._stderr_task is not None:
            self._stderr_task.cancel()
            await asyncio.gather(self._stderr_task, return_exceptions=True)

    @property
    def closed(self) -> bool:
        return self._closed
