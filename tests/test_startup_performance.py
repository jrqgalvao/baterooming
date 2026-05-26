import subprocess
import sys
import textwrap
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class StartupPerformanceTests(unittest.TestCase):
    def test_app_import_does_not_load_excel_processing_cores(self):
        script = textwrap.dedent(
            """
            import sys
            import types

            sys.modules["webview"] = types.SimpleNamespace()
            import app

            loaded = {
                name: name in sys.modules
                for name in ("bate_rooming_core", "match_nomes_core", "pandas", "openpyxl")
            }
            print(loaded)
            """
        )

        result = subprocess.run(
            [sys.executable, "-c", script],
            cwd=ROOT,
            check=True,
            capture_output=True,
            text=True,
        )

        self.assertEqual(
            result.stdout.strip(),
            "{'bate_rooming_core': False, 'match_nomes_core': False, 'pandas': False, 'openpyxl': False}",
        )

    def test_pyinstaller_spec_uses_onedir_without_upx(self):
        spec = (ROOT / "app.spec").read_text(encoding="utf-8")

        self.assertIn("exclude_binaries=True", spec)
        self.assertIn("COLLECT(", spec)
        self.assertIn("upx=False", spec)


if __name__ == "__main__":
    unittest.main()
