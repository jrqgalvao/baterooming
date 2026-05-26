import unittest
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class BateRoomingIgnoreRoomUITests(unittest.TestCase):
    def test_ui_has_checkbox_and_sends_ignore_room_option(self):
        html = (ROOT / "bate_rooming_ui.html").read_text(encoding="utf-8")

        self.assertIn('id="ignore-room-checkbox"', html)
        option = re.search(
            r'<label class="checkbox-option" for="ignore-room-checkbox">(.*?)</label>',
            html,
            re.S,
        )
        self.assertIsNotNone(option)
        visible_text = re.sub(r"<[^>]+>", " ", option.group(1))
        visible_text = " ".join(visible_text.split())
        self.assertEqual(visible_text, "IGNORAR QUARTO")
        self.assertIn("const ignorarQuarto = document.getElementById(\"ignore-room-checkbox\").checked", html)
        self.assertIn("window.pywebview.api.executar({ ignorar_quarto: ignorarQuarto })", html)

    def test_ui_hides_room_columns_and_keeps_status_ok_in_ignore_mode(self):
        html = (ROOT / "bate_rooming_ui.html").read_text(encoding="utf-8")

        self.assertIn('class="room-col"', html)
        self.assertIn("setRoomColumnsHidden(ignorarQuarto)", html)
        self.assertIn("table.classList.toggle(\"hide-room-cols\", hidden)", html)
        self.assertIn("if (!parts.length) return '<span class=\"pill pill-ok\">OK</span>'", html)
        self.assertNotIn("OK · quarto ignorado", html)


if __name__ == "__main__":
    unittest.main()
