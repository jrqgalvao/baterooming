import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class MenuLoaderUiTests(unittest.TestCase):
    def test_menu_has_startup_loader_animation(self):
        html = (ROOT / "menu_ui.html").read_text(encoding="utf-8")

        self.assertIn('class="startup-loader"', html)
        self.assertIn('src="placeholder_logo.svg"', html)
        self.assertIn("hideStartupLoader", html)
        self.assertIn("body.loader-done .startup-loader", html)
        self.assertIn("@media (prefers-reduced-motion: reduce)", html)

    def test_color_logo_is_packaged_for_loader(self):
        spec = (ROOT / "app.spec").read_text(encoding="utf-8")

        self.assertIn("placeholder_logo.svg", spec)


if __name__ == "__main__":
    unittest.main()
