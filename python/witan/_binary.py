from __future__ import annotations

import os
import shutil
import stat
import subprocess
import sys
from pathlib import Path


def _candidate_binary_names() -> tuple[str, ...]:
    return ("witan.exe", "witan") if sys.platform == "win32" else ("witan", "witan.exe")


def _ensure_executable(path: Path) -> Path:
    if sys.platform != "win32":
        mode = path.stat().st_mode
        if not (mode & stat.S_IXUSR):
            path.chmod(mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
    return path


def get_binary_path() -> str:
    """Return the Witan CLI binary path used by the Python SDK."""

    env_path = os.environ.get("WITAN_BINARY")
    if env_path:
        path = Path(env_path).expanduser()
        if path.exists():
            return str(_ensure_executable(path))

    package_dir = Path(__file__).resolve().parent
    for name in _candidate_binary_names():
        bundled = package_dir / "bin" / name
        if bundled.exists():
            return str(_ensure_executable(bundled))

    repo_binary = package_dir.parents[1] / ("witan.exe" if sys.platform == "win32" else "witan")
    if repo_binary.exists():
        return str(_ensure_executable(repo_binary))

    found = shutil.which("witan")
    if found:
        return found

    raise FileNotFoundError(
        "witan binary not found; install the wheel with bundled binary, build ./witan, "
        "set WITAN_BINARY, or put witan on PATH"
    )


def main() -> None:
    """Execute the bundled or discoverable Witan CLI."""

    binary = get_binary_path()
    argv = [binary, *sys.argv[1:]]
    if sys.platform == "win32":
        raise SystemExit(subprocess.call(argv))
    os.execvp(binary, argv)
