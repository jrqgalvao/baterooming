import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class UIClearSingleFileTests(unittest.TestCase):
    def test_bate_rooming_x_button_uses_single_file_clear_api(self):
        html = (ROOT / "bate_rooming_ui.html").read_text(encoding="utf-8")

        self.assertIn("async function limparArquivo(evt, num)", html)
        self.assertIn("window.pywebview.api.limpar_arquivo(num)", html)
        self.assertNotIn("function limparArquivo(evt, num) {\n  evt.preventDefault();\n  evt.stopPropagation();\n  limparTudo();", html)

    def test_match_nomes_x_button_uses_single_file_clear_api(self):
        html = (ROOT / "match_nomes_ui.html").read_text(encoding="utf-8")

        self.assertIn("async function limparArquivo(evt, num)", html)
        self.assertIn("window.pywebview.api.mn_limpar_arquivo(num)", html)
        self.assertNotIn("function limparArquivo(evt, num) {\n  evt.stopPropagation();\n  if (_isBusy) return;\n  limparTudo();", html)


if __name__ == "__main__":
    unittest.main()
