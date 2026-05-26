import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class MenuVersionUiTests(unittest.TestCase):
    def test_menu_shows_current_version(self):
        html = (ROOT / "menu_ui.html").read_text(encoding="utf-8")

        self.assertIn('<span class="version" aria-label="Versão 4.4">v4.4</span>', html)
        self.assertIn('<span class="footer-author">by Rafael Queiroz</span>', html)
        self.assertNotIn('id="changelog-modal"', html)
        self.assertNotIn("function openChangelog()", html)
        self.assertNotIn("function closeChangelog(", html)
        self.assertNotIn("onclick=\"openChangelog()\"", html)
        self.assertNotIn('class="version">v4.0', html)


if __name__ == "__main__":
    unittest.main()
