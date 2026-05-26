import unittest
from unittest.mock import patch

from bate_rooming_core import processar_arquivos
from match_nomes_core import executar_match


class _FailingWorksheet:
    @property
    def max_row(self):
        raise RuntimeError("forced worksheet failure")


class _Workbook:
    def __init__(self, active=None):
        self.active = active if active is not None else object()
        self.closed = False

    def close(self):
        self.closed = True


class WorkbookLifecycleTests(unittest.TestCase):
    def test_match_reader_closes_workbook_when_read_fails(self):
        workbook = _Workbook(active=_FailingWorksheet())

        with patch("match_nomes_core.load_workbook", return_value=workbook):
            result = executar_match("planilha_1.xlsx", "planilha_2.xlsx", 65)

        self.assertFalse(result["ok"])
        self.assertTrue(workbook.closed)

    def test_bate_rooming_closes_loaded_workbooks_when_read_fails(self):
        workbook_1 = _Workbook()
        workbook_2 = _Workbook()

        with patch("bate_rooming_core.ensure_xlsx", side_effect=lambda path: path), \
             patch("bate_rooming_core.openpyxl.load_workbook", side_effect=[workbook_1, workbook_2]), \
             patch("bate_rooming_core.read_sheet", side_effect=ValueError("forced read failure")):
            with self.assertRaises(ValueError):
                processar_arquivos("planilha_1.xlsx", "planilha_2.xlsx")

        self.assertTrue(workbook_1.closed)
        self.assertTrue(workbook_2.closed)


if __name__ == "__main__":
    unittest.main()
