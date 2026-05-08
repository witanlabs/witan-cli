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
)
from .types import Regex
from .workbook import AsyncWorkbook, Workbook

try:
    __version__ = version("witan")
except PackageNotFoundError:
    __version__ = os.environ.get("WITAN_PY_VERSION", "0.0.0")

__all__ = [
    "AsyncWorkbook",
    "Regex",
    "Workbook",
    "WitanError",
    "WitanProcessError",
    "WitanRPCError",
    "WitanTimeoutError",
    "get_binary_path",
    "main",
]
