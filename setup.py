from __future__ import annotations

import os

from setuptools import setup
from setuptools.command.bdist_wheel import bdist_wheel as _bdist_wheel


class bdist_wheel(_bdist_wheel):
    def finalize_options(self) -> None:
        plat_name = os.environ.get("WITAN_WHEEL_PLAT_NAME")
        if plat_name:
            self.plat_name = plat_name
        super().finalize_options()
        self.root_is_pure = False
        if plat_name:
            self.plat_name = plat_name
            self.plat_name_supplied = True

    def get_tag(self) -> tuple[str, str, str]:
        _python, _abi, plat = super().get_tag()
        return "py3", "none", plat


setup(
    version=os.environ.get("WITAN_PY_VERSION", "0.0.0"),
    cmdclass={"bdist_wheel": bdist_wheel},
)
