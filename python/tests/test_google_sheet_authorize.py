from __future__ import annotations

import asyncio
import stat
from pathlib import Path

import pytest

from witan import (
    AsyncGoogleSheet,
    GoogleSheet,
    WitanProcessError,
    WitanRPCError,
    is_google_auth_required,
    is_needs_file_authorization,
)


def fake_binary() -> Path:
    path = Path(__file__).with_name("fake_witan_rpc.py")
    path.chmod(path.stat().st_mode | stat.S_IXUSR)
    return path


def test_authorize_url_when_not_authorized() -> None:
    info = GoogleSheet.authorize_url(
        "gs://abc", binary=fake_binary(), env={"WITAN_FAKE_AUTH_MODE": "not_authorized"}
    )
    assert info["authorized"] is False
    assert info["picker_url"].startswith("https://picker")
    assert info["expires_in_seconds"] == 600


def test_authorize_url_when_already_authorized() -> None:
    info = GoogleSheet.authorize_url(
        "gs://abc", binary=fake_binary(), env={"WITAN_FAKE_AUTH_MODE": "authorized"}
    )
    assert info["authorized"] is True


def test_is_authorized() -> None:
    assert (
        GoogleSheet.is_authorized(
            "gs://abc", binary=fake_binary(), env={"WITAN_FAKE_AUTH_MODE": "authorized"}
        )
        is True
    )
    assert (
        GoogleSheet.is_authorized(
            "gs://abc", binary=fake_binary(), env={"WITAN_FAKE_AUTH_MODE": "not_authorized"}
        )
        is False
    )


def test_wait_until_authorized() -> None:
    assert (
        GoogleSheet.wait_until_authorized(
            "gs://abc", binary=fake_binary(), env={"WITAN_FAKE_AUTH_MODE": "authorized"}
        )
        is True
    )


def test_async_authorize_url() -> None:
    async def run() -> None:
        info = await AsyncGoogleSheet.authorize_url(
            "gs://abc", binary=fake_binary(), env={"WITAN_FAKE_AUTH_MODE": "not_authorized"}
        )
        assert info["picker_url"].startswith("https://picker")

    asyncio.run(run())


def test_is_needs_file_authorization_from_process_error() -> None:
    err = WitanProcessError(
        "witan subprocess exited before responding",
        stderr_tail=("This Google Sheet must be authorized before Witan can access it.",),
    )
    assert is_needs_file_authorization(err)


def test_is_needs_file_authorization_from_rpc_error() -> None:
    err = WitanRPCError(
        "nope", method="m", op="o", request_id="1", code="needs_file_authorization"
    )
    assert is_needs_file_authorization(err)


def test_is_needs_file_authorization_false_for_other() -> None:
    assert not is_needs_file_authorization(WitanProcessError("some other failure"))
    assert not is_needs_file_authorization(ValueError("x"))


def test_open_empty_ref_raises_instead_of_creating() -> None:
    # An empty ref (e.g. unset env var) must not silently create a sheet.
    with pytest.raises(ValueError, match="ref is required"):
        GoogleSheet("")
    with pytest.raises(ValueError, match="ref is required"):
        AsyncGoogleSheet("")


def test_create_still_accepts_empty_ref() -> None:
    # GoogleSheet.create() omits the ref; explicit create mode must still allow it,
    # while open (create=False) must not infer create from an empty ref.
    from witan.google_sheet import _resolve_create, _validate_ref_create

    assert _resolve_create("", True) is True
    _validate_ref_create("", True)  # must not raise
    assert _resolve_create("", False) is False
    assert _resolve_create("new", False) is True  # explicit sentinel still works


def test_is_google_auth_required_for_not_connected() -> None:
    # authorize/status before connect fails with the not-connected message;
    # is_google_auth_required must recognize it as "run connect".
    err = WitanProcessError(
        "witan gsheets authorize failed (exit 1)",
        stderr_tail=("Google Sheets is not connected. Run 'witan gsheets connect' first.",),
    )
    assert is_google_auth_required(err)


def test_is_google_auth_required_for_expired_or_revoked() -> None:
    # authorize/status with an expired/revoked connection emits a different
    # phrasing but the same 'witan gsheets connect' remediation.
    err = WitanProcessError(
        "witan gsheets authorize failed (exit 1)",
        stderr_tail=("Google authorization expired or was revoked. Run 'witan gsheets connect' to reconnect.",),
    )
    assert is_google_auth_required(err)
