import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
NAV_PAGES = ("menu_ui.html", "bate_rooming_ui.html", "match_nomes_ui.html")


class NavigationLoaderUiTests(unittest.TestCase):
    def test_all_pages_use_same_delayed_navigation_loader(self):
        for page in NAV_PAGES:
            with self.subTest(page=page):
                html = (ROOT / page).read_text(encoding="utf-8")

                self.assertIn('class="nav-cover"', html)
                self.assertIn('class="nav-loader-stage"', html)
                self.assertIn('src="placeholder_logo.svg"', html)
                self.assertIn("NAV_LOADER_DELAY_MS = 950", html)
                self.assertIn("startNavigationFeedback", html)
                self.assertIn("stopNavigationFeedback", html)
                self.assertIn("@media (prefers-reduced-motion: reduce)", html)

    def test_navigation_starts_before_delayed_loader_is_shown(self):
        expectations = {
            "menu_ui.html": "const feedback = startNavigationFeedback();\n  try {\n    const res = await window.pywebview.api.abrir_ferramenta(tool);",
            "bate_rooming_ui.html": "const feedback = startNavigationFeedback();\n  try {\n    const res = await window.pywebview.api.voltar_menu();",
            "match_nomes_ui.html": "const feedback = startNavigationFeedback();\n  try {\n    const res = await window.pywebview.api.voltar_menu();",
        }

        for page, expected in expectations.items():
            with self.subTest(page=page):
                html = (ROOT / page).read_text(encoding="utf-8")

                self.assertIn(expected, html)
                self.assertNotIn("await waitForNavigationCover();", html)


if __name__ == "__main__":
    unittest.main()
