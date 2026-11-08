"""Template discovery must be explicit and independent of an author's machine."""

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from template_catalog import roots_from_args  # noqa: E402


class TemplateRootsTest(unittest.TestCase):
    def test_explicit_roots_override_environment(self):
        with tempfile.TemporaryDirectory() as root:
            with patch.dict(os.environ, {"PPT_TEMPLATE_ROOTS": "not-a-real-directory"}):
                self.assertEqual(roots_from_args([root]), [Path(root)])

    def test_missing_config_never_scans_current_directory(self):
        previous = Path.cwd()
        with tempfile.TemporaryDirectory() as root:
            try:
                os.chdir(root)
                with patch.dict(os.environ, {"PPT_TEMPLATE_ROOTS": ""}):
                    with self.assertRaisesRegex(ValueError, "No template directory configured"):
                        roots_from_args(None)
                    Path("templates").mkdir()
                    self.assertEqual(roots_from_args(None), [Path("templates")])
            finally:
                os.chdir(previous)

    def test_environment_ignores_empty_entries(self):
        with tempfile.TemporaryDirectory() as root:
            with patch.dict(os.environ, {"PPT_TEMPLATE_ROOTS": os.pathsep + root + os.pathsep}):
                self.assertEqual(roots_from_args(None), [Path(root)])
            with patch.dict(os.environ, {"PPT_TEMPLATE_ROOTS": os.pathsep}):
                with self.assertRaisesRegex(ValueError, "No template directory configured"):
                    roots_from_args(None)


if __name__ == "__main__":
    unittest.main()
