"""Verify optional Office capability and explicit native-validation policies."""

from __future__ import annotations

import builtins
import importlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from pptx import Presentation
from pptx.util import Inches

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))
import build_deck  # noqa: E402
import doctor  # noqa: E402
import native_powerpoint  # noqa: E402
from native_powerpoint import NativePowerPointUnavailable  # noqa: E402


def sample_plan(root: Path) -> dict:
    source = root / "source.pptx"
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    shape.text = "Source text"
    deck.save(source)
    return {"template": "source.pptx", "output": "result.pptx", "require_animation": False,
            "slides": [{"number": 1, "text": {str(shape.shape_id): "Reproduced text"}}]}


class NativePortabilityTest(unittest.TestCase):
    def test_non_windows_has_a_clear_capability_error(self):
        with patch.object(native_powerpoint.sys, "platform", "linux"):
            with self.assertRaisesRegex(NativePowerPointUnavailable, "require Windows"):
                native_powerpoint.create_powerpoint()

    def test_missing_pywin32_has_a_clear_install_message(self):
        with patch.object(native_powerpoint.sys, "platform", "win32"), \
                patch.object(native_powerpoint.importlib, "import_module", side_effect=ImportError("missing")):
            with self.assertRaisesRegex(NativePowerPointUnavailable, "requirements-windows"):
                native_powerpoint.create_powerpoint()

    def test_host_start_error_is_sanitized(self):
        client = Mock()
        client.DispatchEx.side_effect = RuntimeError("failure at /private/local/path")
        with patch.object(native_powerpoint, "load_win32_client", return_value=client):
            with self.assertRaises(NativePowerPointUnavailable) as captured:
                native_powerpoint.create_powerpoint()
        self.assertNotIn("/private/local/path", str(captured.exception))
        self.assertEqual(captured.exception.code, "powerpoint-unavailable")

    def test_probe_cleanup_error_is_a_sanitized_report(self):
        app = Mock()
        app.Version = "16.0"
        app.Quit.side_effect = RuntimeError("failure at /private/local/path")
        with patch.object(native_powerpoint, "create_powerpoint", return_value=app):
            report = native_powerpoint.probe_powerpoint()
        self.assertFalse(report["available"])
        self.assertEqual(report["code"], "powerpoint-cleanup-failed")
        self.assertNotIn("/private/local/path", json.dumps(report))

    def test_auto_can_build_without_office_and_required_cannot(self):
        unavailable = NativePowerPointUnavailable("powerpoint-unavailable", "PowerPoint is unavailable")
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            plan = sample_plan(root)
            with patch.object(build_deck, "create_powerpoint", side_effect=unavailable):
                result = build_deck.build(plan, root)
                self.assertFalse(result["native_validated"])
                self.assertEqual(Presentation(result["output"]).slides[0].shapes[0].text, "Reproduced text")
                good_output = (root / "result.pptx").read_bytes()
                with self.assertRaises(NativePowerPointUnavailable):
                    build_deck.build({**plan, "native_validation": "required"}, root)
                self.assertEqual((root / "result.pptx").read_bytes(), good_output)

    def test_invalid_deck_is_not_treated_as_host_unavailable(self):
        app = Mock()
        app.Presentations.Open.side_effect = ValueError("Invalid PPTX")
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            plan = sample_plan(root)
            output = root / "result.pptx"
            output.write_bytes(b"previous-output")
            with patch.object(build_deck, "create_powerpoint", return_value=app):
                with self.assertRaisesRegex(ValueError, "Invalid PPTX"):
                    build_deck.build(plan, root)
            self.assertEqual(output.read_bytes(), b"previous-output")
            app.Quit.assert_called_once()
            self.assertFalse(list(root.glob("ppt-build-*.pptx")))

    def test_off_policy_does_not_start_office(self):
        with patch.object(build_deck, "create_powerpoint") as start:
            self.assertFalse(build_deck.verify_native_powerpoint(Path("unused.pptx"), 1, "off"))
            start.assert_not_called()

    def test_cli_can_override_required_policy_for_portable_build(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            plan = {**sample_plan(root), "native_validation": "required"}
            path = root / "plan.json"
            path.write_text(json.dumps(plan), encoding="utf-8")
            result = subprocess.run([sys.executable, "-X", "utf8", str(SCRIPTS / "build_deck.py"), str(path),
                                     "--native-validation", "off"], capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertFalse(json.loads(result.stdout)["native_validated"])

    def test_preview_import_does_not_require_com(self):
        real_import = builtins.__import__

        def no_com(name, *args, **kwargs):
            if name.startswith("win32com") or name == "pywintypes":
                raise ImportError("COM is unavailable")
            return real_import(name, *args, **kwargs)

        with patch("builtins.__import__", side_effect=no_com):
            importlib.reload(importlib.import_module("preview_template"))

    def test_doctor_skip_native_does_not_probe(self):
        with patch.object(doctor, "probe_powerpoint") as probe:
            report = doctor.environment_report(skip_native=True)
            probe.assert_not_called()
        self.assertIsNone(report["native"]["available"])
        self.assertFalse(report["native_ready"])
        self.assertNotIn("executable", report["python"])


if __name__ == "__main__":
    unittest.main()
