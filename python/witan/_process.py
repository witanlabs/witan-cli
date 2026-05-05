from __future__ import annotations

import json
import os
import queue
import subprocess
import threading
from collections import deque
from collections.abc import Mapping
from typing import Any

from .exceptions import WitanProcessError, WitanRPCError, WitanTimeoutError


class StdioRPCProcess:
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
        self._timeout = 90.0 if request_timeout is None else request_timeout
        self._lock = threading.Lock()
        self._responses: queue.Queue[str | None] = queue.Queue()
        self._stderr: deque[str] = deque(maxlen=50)
        self._closed = False
        try:
            self._proc = subprocess.Popen(
                argv,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                bufsize=0,
                env=merged_env,
            )
        except OSError as exc:
            raise WitanProcessError(f"starting witan subprocess: {exc}") from exc
        if self._proc.stdin is None or self._proc.stdout is None or self._proc.stderr is None:
            self.terminate()
            raise WitanProcessError("starting witan subprocess: stdio pipes unavailable")
        self._stdout_thread = threading.Thread(target=self._drain_stdout, daemon=True)
        self._stderr_thread = threading.Thread(target=self._drain_stderr, daemon=True)
        self._stdout_thread.start()
        self._stderr_thread.start()

    def _drain_stdout(self) -> None:
        assert self._proc.stdout is not None
        for line in self._proc.stdout:
            text = line.decode("utf-8", errors="replace").rstrip("\r\n")
            self._responses.put(text)
        self._responses.put(None)

    def _drain_stderr(self) -> None:
        assert self._proc.stderr is not None
        for line in self._proc.stderr:
            text = line.decode("utf-8", errors="replace").rstrip("\r\n")
            if text:
                self._stderr.append(text)

    def _stderr_tail(self) -> tuple[str, ...]:
        return tuple(self._stderr)

    def request(self, *, request_id: str, op: str, args: Mapping[str, Any], method: str) -> Any:
        with self._lock:
            if self._closed:
                raise WitanProcessError("witan subprocess is closed", stderr_tail=self._stderr_tail())
            payload = json.dumps(
                {"id": request_id, "op": op, "args": dict(args)},
                separators=(",", ":"),
            )
            assert self._proc.stdin is not None
            try:
                self._proc.stdin.write((payload + "\n").encode("utf-8"))
                self._proc.stdin.flush()
            except OSError as exc:
                self.terminate()
                raise WitanProcessError(f"writing RPC request: {exc}", stderr_tail=self._stderr_tail()) from exc

            try:
                line = self._responses.get(timeout=self._timeout)
            except queue.Empty as exc:
                self.terminate()
                raise WitanTimeoutError(
                    f"RPC timeout: {method} ({op}) did not respond within {self._timeout:g}s",
                    stderr_tail=self._stderr_tail(),
                ) from exc

            if line is None:
                code = self._proc.poll()
                raise WitanProcessError(
                    f"witan subprocess exited before responding (exit={code})",
                    stderr_tail=self._stderr_tail(),
                )

            try:
                response = json.loads(line)
            except json.JSONDecodeError as exc:
                self.terminate()
                raise WitanProcessError(f"invalid JSON RPC response: {line!r}", stderr_tail=self._stderr_tail()) from exc

            if not isinstance(response, dict):
                self.terminate()
                raise WitanProcessError(f"invalid RPC response object: {response!r}", stderr_tail=self._stderr_tail())

            response_id = response.get("id")
            if response_id not in (None, "", request_id):
                self.terminate()
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
                self.terminate()
                raise WitanProcessError(f"invalid RPC ok field: {response!r}", stderr_tail=self._stderr_tail())
            return response.get("result")

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        if self._proc.stdin is not None:
            try:
                self._proc.stdin.close()
            except OSError:
                pass
        try:
            self._proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            self.terminate()

    def terminate(self) -> None:
        self._closed = True
        if self._proc.poll() is None:
            self._proc.terminate()
            try:
                self._proc.wait(timeout=2)
            except subprocess.TimeoutExpired:
                self._proc.kill()
                self._proc.wait(timeout=2)

    @property
    def closed(self) -> bool:
        return self._closed
