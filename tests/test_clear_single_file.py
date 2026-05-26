import sys
import types
import unittest


sys.modules.setdefault("webview", types.SimpleNamespace())
from app import AppAPI


class ClearSingleFileTests(unittest.TestCase):
    def test_bate_rooming_clear_file_removes_only_selected_path(self):
        api = AppAPI()
        api._br_path1 = "planilha_1.xlsx"
        api._br_path2 = "planilha_2.xlsx"
        api._br_results = [{"nome": "resultado"}]

        result = api.limpar_arquivo(1)

        self.assertEqual(result, {"ok": True})
        self.assertEqual(api._br_path1, "")
        self.assertEqual(api._br_path2, "planilha_2.xlsx")
        self.assertEqual(api._br_results, [])

    def test_match_nomes_clear_file_removes_only_selected_path(self):
        api = AppAPI()
        api._mn_path1 = "planilha_1.xlsx"
        api._mn_path2 = "planilha_2.xlsx"
        api._mn_nomes_finais = ["Nome"]
        api._mn_statuses = ["Match encontrado"]
        api._mn_scores = [100.0]

        result = api.mn_limpar_arquivo(2)

        self.assertEqual(result, {"ok": True})
        self.assertEqual(api._mn_path1, "planilha_1.xlsx")
        self.assertEqual(api._mn_path2, "")
        self.assertEqual(api._mn_nomes_finais, [])
        self.assertEqual(api._mn_statuses, [])
        self.assertEqual(api._mn_scores, [])


if __name__ == "__main__":
    unittest.main()
