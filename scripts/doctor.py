"""Report portable dependencies and optional PowerPoint capability as JSON."""

from __future__ import annotations

import argparse
import importlib
import importlib.metadata
import json
import platform
import sys

from native_powerpoint import probe_powerpoint


DEPENDENCIES = {"python-pptx": "pptx", "Pillow": "PIL", "lxml": "lxml.etree",
                "XlsxWriter": "xlsxwriter", "typing_extensions": "typing_extensions"}


def environment_report(skip_native: bool = False) -> dict:
    dependencies = {}
    for package, module in DEPENDENCIES.items():
        try:
            version = importlib.metadata.version(package)
        except importlib.metadata.PackageNotFoundError:
            version = None
        try:
            importlib.import_module(module)
            status = "ready"
        except (ImportError, OSError):
            status = "unavailable"
        dependencies[package] = {"status": status, "version": version}
    python_supported = (3, 10) <= sys.version_info[:2] <= (3, 13)
    portable_ready = python_supported and all(item["status"] == "ready" for item in dependencies.values())
    native = ({"available": None, "code": "not-probed", "message": "Native host probe was explicitly skipped."}
              if skip_native else probe_powerpoint())
    return {
        "python": {"version": platform.python_version(), "supported": python_supported},
        "platform": {"system": platform.system(), "machine": platform.machine()},
        "dependencies": dependencies,
        "portable_ready": portable_ready,
        "native": native,
        "native_ready": portable_ready and native["available"] is True,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--require-native", action="store_true", help="Exit nonzero unless native PowerPoint operations can run")
    group.add_argument("--skip-native", action="store_true", help="Check portable dependencies without starting PowerPoint")
    args = parser.parse_args()
    report = environment_report(args.skip_native)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["portable_ready"] and (not args.require_native or report["native_ready"]) else 1


if __name__ == "__main__":
    raise SystemExit(main())
