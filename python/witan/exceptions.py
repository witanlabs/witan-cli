from __future__ import annotations

from typing import Any, Mapping


class WitanError(Exception):
    """Base exception for the Witan Python SDK."""


class WitanProcessError(WitanError):
    """Raised when the underlying witan subprocess cannot complete a request."""

    def __init__(self, message: str, *, stderr_tail: tuple[str, ...] = ()) -> None:
        if stderr_tail:
            message = f"{message}\nstderr tail:\n" + "\n".join(stderr_tail)
        super().__init__(message)
        self.stderr_tail = stderr_tail


class WitanTimeoutError(WitanProcessError):
    """Raised when a workbook RPC request exceeds its timeout."""


class WitanRPCError(WitanError):
    """Raised when the xlsx RPC endpoint returns ok=false."""

    def __init__(
        self,
        message: str,
        *,
        method: str,
        op: str,
        request_id: str,
        code: str | None = None,
        response: Mapping[str, Any] | None = None,
    ) -> None:
        label = code or "RPC_ERROR"
        super().__init__(f"{label}: {message}")
        self.method = method
        self.op = op
        self.request_id = request_id
        self.code = code
        self.response = dict(response or {})
