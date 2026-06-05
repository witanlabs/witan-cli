"""Python SDK and CLI entry point for Witan."""

from __future__ import annotations

import os
from importlib.metadata import PackageNotFoundError, version

from ._binary import get_binary_path, main
from .exceptions import (
    WitanError,
    WitanProcessError,
    WitanRPCError,
    WitanTimeoutError,
    is_google_auth_required,
)
from .google_sheet import AsyncGoogleSheet, GoogleSheet
from ._spreadsheet_base import Regex
from .workbook import AsyncWorkbook, Workbook

try:
    __version__ = version("witan")
except PackageNotFoundError:
    __version__ = os.environ.get("WITAN_PY_VERSION", "0.0.0")

__all__ = [
    "AsyncGoogleSheet",
    "AsyncWorkbook",
    "GoogleSheet",
    "Regex",
    "Workbook",
    "WitanError",
    "WitanProcessError",
    "WitanRPCError",
    "WitanTimeoutError",
    "get_binary_path",
    "is_google_auth_required",
    "main",
]
