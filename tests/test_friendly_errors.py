import sys
import tempfile
import types
import unittest
import importlib.util
from pathlib import Path


sys.modules.setdefault("webview", types.SimpleNamespace())

import app
from bate_rooming_core import ensure_xlsx
from match_nomes_core import executar_match


ROOT = Path(__file__).resolve().parents[1]


class FriendlyErrorTests(unittest.TestCase):
    def test_friendly_error_hides_python_call_signature_details(self):
        message = app._friendly_error(
            TypeError("AppAPI.executar() takes 1 positional argument but 2 were given"),
            "Bate-Rooming",
        )

        self.assertNotIn("TypeError", message)
        self.assertNotIn("AppAPI", message)
        self.assertNotIn("positional argument", message)
        self.assertIn("Não conseguimos continuar", message)
        self.assertIn("tente novamente", message.lower())

    def test_friendly_error_hides_tracebacks(self):
        message = app._friendly_error(
            ValueError("Traceback (most recent call last):\nFile x.py\nValueError: bad"),
            "Match de Nomes",
        )

        self.assertNotIn("Traceback", message)
        self.assertNotIn("ValueError", message)
        self.assertIn("Não conseguimos continuar", message)

    def test_bate_rooming_ui_does_not_show_raw_unexpected_errors(self):
        html = (ROOT / "bate_rooming_ui.html").read_text(encoding="utf-8")

        self.assertIn("friendlyClientError", html)
        self.assertNotIn('showToast("Erro: " + e', html)
        self.assertNotIn('showToast("Erro ao exportar: " + e', html)

    def test_match_nomes_ui_does_not_show_raw_unexpected_errors(self):
        html = (ROOT / "match_nomes_ui.html").read_text(encoding="utf-8")

        self.assertIn("friendlyClientError", html)
        self.assertNotIn('showToast("Erro: " + String(e)', html)
        self.assertNotIn('showToast("Erro ao exportar: " + e', html)

    def test_match_nomes_core_does_not_return_raw_traceback(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad_file = Path(tmp) / "arquivo_ruim.xlsx"
            bad_file.write_text("nao e uma planilha valida", encoding="utf-8")

            result = executar_match(str(bad_file), str(bad_file), 65)

        self.assertFalse(result["ok"])
        self.assertNotIn("Traceback", result["erro"])
        self.assertNotIn("File \"", result["erro"])

    @unittest.skipIf(importlib.util.find_spec("xlrd") is not None, "xlrd installed")
    def test_bate_rooming_xls_missing_dependency_has_clear_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_excel = Path(tmp) / "antigo.xls"
            old_excel.write_text("conteudo irrelevante", encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "requirements.txt|xlsx"):
                ensure_xlsx(str(old_excel))


if __name__ == "__main__":
    unittest.main()
