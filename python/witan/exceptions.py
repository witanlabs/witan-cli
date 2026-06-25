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
    """Raised when an RPC endpoint returns ok=false."""

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


_GOOGLE_AUTH_REQUIRED_MARKERS = (
    # Error codes (when surfaced as a code or in JSON).
    "google_auth_required",
    "google_sheets_not_connected",
    "google_sheets_scope_not_granted",
    # Every "needs connect/reconnect" CLI message ends with this remediation
    # (not-connected, expired/revoked, and sheet-op google_auth_required alike),
    # so match the command rather than each individual phrasing.
    "witan gsheets connect",
)


def _text_indicates_google_auth_required(text: str) -> bool:
    return any(marker in text for marker in _GOOGLE_AUTH_REQUIRED_MARKERS)


def is_google_auth_required(err: BaseException) -> bool:
    """Return True when *err* indicates the Google account must be connected or
    re-authorized — i.e. the caller should run ``witan gsheets connect``.

    Covers both an expired/revoked connection and the not-yet-connected case
    (which surfaces from the authorize-sheet / status path, e.g. when
    :meth:`GoogleSheet.authorize_url` is called before connecting)."""
    if isinstance(err, WitanRPCError) and err.code in (
        "google_auth_required",
        "google_sheets_not_connected",
        "google_sheets_scope_not_granted",
    ):
        return True
    if isinstance(err, WitanProcessError):
        if _text_indicates_google_auth_required(str(err)):
            return True
        return any(_text_indicates_google_auth_required(line) for line in err.stderr_tail)
    return False


_NEEDS_FILE_AUTHORIZATION_MARKERS = (
    "needs_file_authorization",
    "must be authorized before Witan",
)


def _text_indicates_needs_file_authorization(text: str) -> bool:
    return any(marker in text for marker in _NEEDS_FILE_AUTHORIZATION_MARKERS)


def is_needs_file_authorization(err: BaseException) -> bool:
    """Return True when *err* indicates the specific spreadsheet has not been
    authorized for the app (drive.file scope).

    Recover by authorizing the sheet and retrying::

        try:
            with GoogleSheet.open(ref) as sheet:
                ...
        except Exception as err:
            if is_needs_file_authorization(err):
                info = GoogleSheet.authorize_url(ref)
                hand_url_to_human(info["picker_url"])  # opens Google's picker
                GoogleSheet.wait_until_authorized(ref)
                # retry GoogleSheet.open(ref)
    """
    if isinstance(err, WitanRPCError) and err.code == "needs_file_authorization":
        return True
    if isinstance(err, WitanProcessError):
        if _text_indicates_needs_file_authorization(str(err)):
            return True
        return any(_text_indicates_needs_file_authorization(line) for line in err.stderr_tail)
    return False
