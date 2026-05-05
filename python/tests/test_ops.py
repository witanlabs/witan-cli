from __future__ import annotations

import inspect
import re

from witan import AsyncWorkbook, Workbook
from witan._ops import OPERATIONS


def public_methods(cls: type[object]) -> set[str]:
    return {
        name
        for name, value in inspect.getmembers(cls)
        if callable(value) and not name.startswith("_")
    }


def test_operation_table_methods_exist_on_sync_and_async_workbooks() -> None:
    sync_methods = public_methods(Workbook)
    async_methods = public_methods(AsyncWorkbook)
    for operation in OPERATIONS:
        assert operation.method in sync_methods
        assert operation.method in async_methods


def test_public_workbook_methods_are_snake_case_only() -> None:
    pattern = re.compile(r"^[a-z][a-z0-9_]*$")
    for cls in (Workbook, AsyncWorkbook):
        methods = public_methods(cls)
        assert "call" not in methods
        assert all(pattern.match(name) for name in methods)
