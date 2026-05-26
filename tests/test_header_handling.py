import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from bate_rooming_core import Status, processar_arquivos
from match_nomes_core import executar_match


def _write_rows(path: Path, rows: list[list[object]]) -> None:
    wb = Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)


class HeaderHandlingTests(unittest.TestCase):
    def test_match_accepts_name_column_when_rooming_headers_are_present(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            planilha_1 = base / "match_header_1.xlsx"
            planilha_2 = base / "match_header_2.xlsx"

            _write_rows(
                planilha_1,
                [
                    ["quarto", "nome", "check in", "check out"],
                    ["101", "Helton Machado Kraus", "01/01/2026", "02/01/2026"],
                ],
            )
            _write_rows(
                planilha_2,
                [
                    ["quarto", "nome", "check in", "check out"],
                    ["201", "HELTON KRAUS", "01/01/2026", "02/01/2026"],
                    ["202", "LAIS MACHADO KLEIN", "01/01/2026", "02/01/2026"],
                ],
            )

            result = executar_match(str(planilha_1), str(planilha_2), 65)

        self.assertTrue(result["ok"], result.get("erro"))
        self.assertEqual(
            result["nomes_finais"],
            ["Helton Machado Kraus", "LAIS MACHADO KLEIN"],
        )
        self.assertEqual(result["statuses"], ["Match encontrado", "Nao encontrado"])

    def test_match_no_header_input_still_uses_first_column(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            planilha_1 = base / "match_no_header_1.xlsx"
            planilha_2 = base / "match_no_header_2.xlsx"

            _write_rows(planilha_1, [["Helton Machado Kraus"]])
            _write_rows(planilha_2, [["HELTON KRAUS"], ["LAIS MACHADO KLEIN"]])

            result = executar_match(str(planilha_1), str(planilha_2), 65)

        self.assertTrue(result["ok"], result.get("erro"))
        self.assertEqual(
            result["nomes_finais"],
            ["Helton Machado Kraus", "LAIS MACHADO KLEIN"],
        )
        self.assertEqual(result["statuses"], ["Match encontrado", "Nao encontrado"])

    def test_match_detects_name_header_in_first_column_with_extra_columns(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            planilha_1 = base / "match_header_col_a_1.xlsx"
            planilha_2 = base / "match_header_col_a_2.xlsx"

            _write_rows(
                planilha_1,
                [
                    ["nome", "email"],
                    ["Ana Maria Oliveira", "ana@planilha1.test"],
                ],
            )
            _write_rows(
                planilha_2,
                [
                    ["nome", "email"],
                    ["ANA MARIA OLIVEIRA", "ana@planilha2.test"],
                    ["BRUNO COSTA", "bruno@planilha2.test"],
                ],
            )

            result = executar_match(str(planilha_1), str(planilha_2), 65)

        self.assertTrue(result["ok"], result.get("erro"))
        self.assertEqual(result["nomes_finais"], ["Ana Maria Oliveira", "BRUNO COSTA"])
        self.assertEqual(result["statuses"], ["Match encontrado", "Nao encontrado"])
        self.assertEqual(result["name_column"], 1)

    def test_match_no_header_two_column_list_uses_first_column(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            planilha_1 = base / "match_no_header_two_col_1.xlsx"
            planilha_2 = base / "match_no_header_two_col_2.xlsx"

            _write_rows(planilha_1, [["Ana Maria Oliveira", "ana@planilha1.test"]])
            _write_rows(
                planilha_2,
                [
                    ["ANA MARIA OLIVEIRA", "ana@planilha2.test"],
                    ["BRUNO COSTA", "bruno@planilha2.test"],
                ],
            )

            result = executar_match(str(planilha_1), str(planilha_2), 65)

        self.assertTrue(result["ok"], result.get("erro"))
        self.assertEqual(result["nomes_finais"], ["Ana Maria Oliveira", "BRUNO COSTA"])
        self.assertEqual(result["statuses"], ["Match encontrado", "Nao encontrado"])
        self.assertEqual(result["name_column"], 1)

    def test_match_no_header_rooming_layout_uses_name_column(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            planilha_1 = base / "match_rooming_no_header_1.xlsx"
            planilha_2 = base / "match_rooming_no_header_2.xlsx"

            _write_rows(
                planilha_1,
                [["101", "Helton Machado Kraus", "01/01/2026", "02/01/2026"]],
            )
            _write_rows(
                planilha_2,
                [
                    ["201", "HELTON KRAUS", "01/01/2026", "02/01/2026"],
                    ["202", "LAIS MACHADO KLEIN", "01/01/2026", "02/01/2026"],
                ],
            )

            result = executar_match(str(planilha_1), str(planilha_2), 65)

        self.assertTrue(result["ok"], result.get("erro"))
        self.assertEqual(
            result["nomes_finais"],
            ["Helton Machado Kraus", "LAIS MACHADO KLEIN"],
        )
        self.assertEqual(result["statuses"], ["Match encontrado", "Nao encontrado"])

    def test_bate_rooming_results_match_with_and_without_headers(self):
        header_rows_1 = [
            ["quarto", "nome", "check in", "check out"],
            ["101", "Helton Machado Kraus", "01/01/2026", "03/01/2026"],
        ]
        header_rows_2 = [
            ["quarto", "nome", "check in", "check out"],
            ["101", "Helton Machado Kraus", "01/01/2026", "03/01/2026"],
        ]
        no_header_rows_1 = [header_rows_1[1]]
        no_header_rows_2 = [header_rows_2[1]]

        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            with_header_1 = base / "bate_header_1.xlsx"
            with_header_2 = base / "bate_header_2.xlsx"
            no_header_1 = base / "bate_no_header_1.xlsx"
            no_header_2 = base / "bate_no_header_2.xlsx"
            _write_rows(with_header_1, header_rows_1)
            _write_rows(with_header_2, header_rows_2)
            _write_rows(no_header_1, no_header_rows_1)
            _write_rows(no_header_2, no_header_rows_2)

            results_header, warnings_header, kpis_header = processar_arquivos(
                str(with_header_1),
                str(with_header_2),
            )
            results_no_header, warnings_no_header, kpis_no_header = processar_arquivos(
                str(no_header_1),
                str(no_header_2),
            )

        self.assertEqual(warnings_header, [])
        self.assertEqual(warnings_no_header, [])
        self.assertEqual(kpis_header, kpis_no_header)
        self.assertEqual(len(results_header), 1)
        self.assertEqual(len(results_no_header), 1)
        self.assertEqual(results_header[0]["nome"], results_no_header[0]["nome"])
        self.assertEqual(results_header[0]["s_quarto"], Status.OK)
        self.assertEqual(results_header[0]["s_ci"], Status.OK)
        self.assertEqual(results_header[0]["s_co"], Status.OK)

    def test_bate_rooming_accepts_headers_in_different_order(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            planilha_1 = base / "bate_shuffled_header_1.xlsx"
            planilha_2 = base / "bate_shuffled_header_2.xlsx"
            rows = [
                ["check out", "nome", "quarto", "check in"],
                ["03/01/2026", "Helton Machado Kraus", "101", "01/01/2026"],
            ]
            _write_rows(planilha_1, rows)
            _write_rows(planilha_2, rows)

            results, warnings, kpis = processar_arquivos(str(planilha_1), str(planilha_2))

        self.assertEqual(warnings, [])
        self.assertEqual(kpis["total"], 1)
        self.assertEqual(results[0]["nome"], "Helton Machado Kraus")
        self.assertEqual(results[0]["s_quarto"], Status.OK)
        self.assertEqual(results[0]["s_ci"], Status.OK)
        self.assertEqual(results[0]["s_co"], Status.OK)

    def test_bate_rooming_accepts_name_only_header_with_missing_field_warnings(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            planilha_1 = base / "bate_name_only_header_1.xlsx"
            planilha_2 = base / "bate_name_only_header_2.xlsx"
            rows = [
                ["codigo interno", "nome", "observacao"],
                ["ABC-123", "Helton Machado Kraus", "VIP"],
            ]
            _write_rows(planilha_1, rows)
            _write_rows(planilha_2, rows)

            results, warnings, kpis = processar_arquivos(str(planilha_1), str(planilha_2))

        self.assertEqual(kpis["total"], 1)
        self.assertEqual(results[0]["nome"], "Helton Machado Kraus")
        self.assertEqual(results[0]["s_quarto"], Status.DATA_AUSENTE)
        self.assertEqual(results[0]["s_ci"], Status.DATA_AUSENTE)
        self.assertEqual(results[0]["s_co"], Status.DATA_AUSENTE)
        self.assertTrue(any("Coluna 'Quarto' não encontrada" in w for w in warnings))
        self.assertTrue(any("Coluna 'Check-in' não encontrada" in w for w in warnings))
        self.assertTrue(any("Coluna 'Check-out' não encontrada" in w for w in warnings))


if __name__ == "__main__":
    unittest.main()
