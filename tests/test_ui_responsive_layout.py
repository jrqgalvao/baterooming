import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class UIResponsiveLayoutTests(unittest.TestCase):
    def test_core_visual_tokens_are_unified_across_pages(self):
        expected_tokens = {
            "--bg:           #F6F7F6;",
            "--surface:      #FFFFFF;",
            "--surface2:     #EEF1EE;",
            "--border:       #C7D3C7;",
            "--text:         #162016;",
            "--text-muted:   #3D4F3D;",
            "--text-faint:   #5F745F;",
            "--primary:      #2E7D32;",
            "--primary-hov:  #1B5E20;",
            "--primary-lt:   #E8F5E9;",
            "--header-bg:    #141E14;",
            "--radius:       8px;",
            "--radius-lg:    8px;",
        }

        for page in ("menu_ui.html", "bate_rooming_ui.html", "match_nomes_ui.html"):
            with self.subTest(page=page):
                html = (ROOT / page).read_text(encoding="utf-8")
                for token in expected_tokens:
                    self.assertIn(token, html)

    def test_tool_pages_have_clear_disabled_button_state(self):
        expected = (
            ".btn:disabled {\n"
            "    background: var(--disabled-bg);\n"
            "    color: var(--disabled-text);\n"
            "    border-color: var(--disabled-border);\n"
            "    box-shadow: none;\n"
            "    opacity: 1;\n"
            "    cursor: not-allowed;\n"
            "    transform: none;\n"
            "  }"
        )

        for page in ("bate_rooming_ui.html", "match_nomes_ui.html"):
            with self.subTest(page=page):
                html = (ROOT / page).read_text(encoding="utf-8")
                self.assertIn("--disabled-bg: #F1F4F1;", html)
                self.assertIn(expected, html)

    def test_menu_desktop_vertical_balance_is_refined(self):
        html = (ROOT / "menu_ui.html").read_text(encoding="utf-8")

        self.assertIn("justify-content: flex-start;", html)
        self.assertIn("padding: clamp(96px, 22vh, 136px) 48px 40px;", html)
        self.assertIn("margin-bottom: 36px;", html)

    def test_menu_mobile_content_is_width_constrained(self):
        html = (ROOT / "menu_ui.html").read_text(encoding="utf-8")

        self.assertIn("overflow-x: hidden;", html)
        self.assertIn("width: min(100%, 480px);", html)
        self.assertIn(".footer-author {\n      min-width: 0;", html)

    def test_bate_rooming_mobile_actions_stack(self):
        html = (ROOT / "bate_rooming_ui.html").read_text(encoding="utf-8")

        self.assertIn("@media (max-width: 480px)", html)
        self.assertIn(".actions-bar {\n    grid-template-columns: 1fr;", html)
        self.assertIn(".actions-bar .btn {\n    width: 100%;", html)

    def test_match_nomes_mobile_actions_stack(self):
        html = (ROOT / "match_nomes_ui.html").read_text(encoding="utf-8")

        self.assertIn("@media (max-width: 480px)", html)
        self.assertIn(".actions-row .btn {\n    flex-basis: 100%;", html)
        self.assertIn("#status-msg { font-size: 12px; color: var(--text-faint); min-width: 0; }", html)

    def test_tool_content_areas_allow_wrapping_on_narrow_windows(self):
        bate_html = (ROOT / "bate_rooming_ui.html").read_text(encoding="utf-8")
        match_html = (ROOT / "match_nomes_ui.html").read_text(encoding="utf-8")

        self.assertRegex(
            bate_html,
            re.compile(r"\.control-sidebar,\s*\.content-area\s*\{[^}]*min-width:\s*0;", re.S),
        )
        self.assertRegex(
            match_html,
            re.compile(r"\.upload-card,\s*\.threshold-card,\s*#result-box\s*\{[^}]*min-width:\s*0;", re.S),
        )


if __name__ == "__main__":
    unittest.main()
