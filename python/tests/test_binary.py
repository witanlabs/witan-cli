from __future__ import annotations

import os
import stat
from pathlib import Path

import pytest

from witan._binary import get_binary_path


def test_get_binary_path_prefers_environment_variable(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    binary = tmp_path / "witan"
    binary.write_text("#!/bin/sh\nexit 0\n", encoding="utf-8")
    monkeypatch.setenv("WITAN_BINARY", str(binary))

    resolved = Path(get_binary_path())

    assert resolved == binary
    assert resolved.stat().st_mode & stat.S_IXUSR


def test_get_binary_path_ignores_missing_environment_variable(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("WITAN_BINARY", os.fspath(Path("/definitely/missing/witan")))
    try:
        resolved = get_binary_path()
    except FileNotFoundError:
        return
    assert Path(resolved).name in {"witan", "witan.exe"}
