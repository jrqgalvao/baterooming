import sys
import types
import unittest


class ExportNameTests(unittest.TestCase):
    def test_match_export_name_uses_planilha_2_stem_plus_corrigido(self):
        sys.modules.setdefault("webview", types.SimpleNamespace())
        from app import _match_export_name

        self.assertEqual(
            _match_export_name(r"C:\dados\rooming_reference.xlsx"),
            "rooming_reference_corrigido.xlsx",
        )
        self.assertEqual(
            _match_export_name(r"C:\dados\Rooming Reference May.xlsx"),
            "Rooming Reference May_corrigido.xlsx",
        )


if __name__ == "__main__":
    unittest.main()
