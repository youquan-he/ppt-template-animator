"""Optional Windows PowerPoint automation without import-time COM dependencies."""

from __future__ import annotations

import importlib
import sys


class NativePowerPointUnavailable(RuntimeError):
    """The native host cannot be started; this does not describe a deck error."""

    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


def load_win32_client():
    if sys.platform != "win32":
        raise NativePowerPointUnavailable(
            "unsupported-platform", "Native PowerPoint operations require Windows with desktop Microsoft PowerPoint. Portable inspection and building remain available.")
    try:
        return importlib.import_module("win32com.client")
    except (ImportError, OSError) as exc:
        raise NativePowerPointUnavailable(
            "pywin32-unavailable", "The optional pywin32 runtime is unavailable. Install requirements-windows.txt in this Python environment.") from exc


def create_powerpoint():
    """Start the native host, converting only startup failures to unavailable."""
    client = load_win32_client()
    try:
        return client.DispatchEx("PowerPoint.Application")
    except Exception as exc:
        # Do not include the raw exception: COM diagnostics can contain local paths.
        code = getattr(exc, "hresult", None)
        detail = f" (HRESULT {code})" if isinstance(code, int) else ""
        raise NativePowerPointUnavailable(
            "powerpoint-unavailable", "Desktop Microsoft PowerPoint could not be started. Install or repair PowerPoint, then run scripts/doctor.py --require-native." + detail) from exc


def powerpoint_constants():
    return load_win32_client().constants


def probe_powerpoint() -> dict:
    """Report host startup capability, not presentation validity or visual quality."""
    try:
        app = create_powerpoint()
    except NativePowerPointUnavailable as exc:
        return {"available": False, "code": exc.code, "message": str(exc)}
    try:
        version = str(app.Version)
    except Exception:
        version = None
    try:
        app.Quit()
    except Exception:
        return {"available": False, "code": "powerpoint-cleanup-failed",
                "message": "PowerPoint started but did not close cleanly. Resolve the Office automation issue before running native tools."}
    return {"available": True, "code": "ready", "version": version,
            "message": "PowerPoint automation started successfully; individual decks still need native validation."}
