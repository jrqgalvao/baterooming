import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook

from core.bate_rooming import Status, normalize_dates, parse_date_raw, processar_arquivos
from core.bate_rooming_export import write_excel as write_bate_excel
from core.match_nomes import executar_match


def _write_rows(path: Path, rows: list[list[object]]) -> None:
    wb = Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


class BusinessRuleCharacterizationTests(unittest.TestCase):
    def test_dates_accept_supported_formats_and_normalize_outlier_year(self):
        self.assertEqual(parse_date_raw("31/12/2026"), date(2026, 12, 31))
        self.assertEqual(parse_date_raw("2026-12-31"), date(2026, 12, 31))
        self.assertIsNone(parse_date_raw("data invalida"))

        normalized = normalize_dates(
            [date(2026, 1, 10), date(2026, 2, 10), date(2036, 3, 10)]
        )

        self.assertEqual(normalized, [date(2026, 1, 10), date(2026, 2, 10), date(2026, 3, 10)])

    def test_bate_rooming_marks_duplicates_and_placeholders_as_repeated(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            sistema = base / "sistema.xlsx"
            hotel = base / "hotel.xlsx"
            _write_rows(
                sistema,
                [
                    ["quarto", "nome", "check in", "check out"],
                    ["101", "Nome Duplicado", "01/01/2026", "02/01/2026"],
                    ["102", "Nome Duplicado", "01/01/2026", "02/01/2026"],
                    ["103", "A nomear", "01/01/2026", "02/01/2026"],
                ],
            )
            _write_rows(
                hotel,
                [["quarto", "nome", "check in", "check out"], ["201", "Outro Nome", "01/01/2026", "02/01/2026"]],
            )

            results, warnings, kpis = processar_arquivos(str(sistema), str(hotel))

        repeated = [row for row in results if row["s_geral"] == Status.REPETIDO]
        self.assertEqual(len(repeated), 3)
        self.assertEqual(kpis["dups"], 3)
        self.assertTrue(any("repetidos/placeholders" in warning for warning in warnings))

    def test_bate_rooming_uses_match_normalization_for_duplicates(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            sistema = base / "sistema.xlsx"
            hotel = base / "hotel.xlsx"
            _write_rows(
                sistema,
                [
                    ["quarto", "nome", "check in", "check out"],
                    ["101", "Nome Duplicado", "01/01/2026", "02/01/2026"],
                    ["102", "Nome-Duplicado", "01/01/2026", "02/01/2026"],
                ],
            )
            _write_rows(
                hotel,
                [
                    ["quarto", "nome", "check in", "check out"],
                    ["201", "NOME DUPLICADO", "01/01/2026", "02/01/2026"],
                ],
            )

            results, warnings, kpis = processar_arquivos(str(sistema), str(hotel))

        self.assertEqual(len(results), 3)
        self.assertTrue(all(row["s_geral"] == Status.REPETIDO for row in results))
        self.assertEqual(kpis["dups"], 3)
        self.assertTrue(any("nome-duplicado" in warning.lower() for warning in warnings))

    def test_bate_rooming_uses_match_normalization_for_placeholders(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            sistema = base / "sistema.xlsx"
            hotel = base / "hotel.xlsx"
            _write_rows(
                sistema,
                [
                    ["quarto", "nome", "check in", "check out"],
                    ["101", "A-Nomear", "01/01/2026", "02/01/2026"],
                ],
            )
            _write_rows(
                hotel,
                [["quarto", "nome", "check in", "check out"], ["201", "Outro Nome", "01/01/2026", "02/01/2026"]],
            )

            results, warnings, kpis = processar_arquivos(str(sistema), str(hotel))

        repeated = [row for row in results if row["s_geral"] == Status.REPETIDO]
        self.assertEqual(len(repeated), 1)
        self.assertEqual(kpis["dups"], 1)
        self.assertTrue(any("a-nomear" in warning.lower() for warning in warnings))

    def test_bate_rooming_warns_when_optional_columns_are_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            sistema = base / "sistema.xlsx"
            hotel = base / "hotel.xlsx"
            _write_rows(sistema, [["nome"], ["Maria Silva"]])
            _write_rows(hotel, [["nome"], ["Maria Silva"]])

            results, warnings, kpis = processar_arquivos(str(sistema), str(hotel))

        self.assertEqual(len(results), 1)
        self.assertEqual(kpis["matched"], 1)
        self.assertEqual(kpis["div"], 1)
        self.assertEqual(len(warnings), 6)
        self.assertTrue(all("tratados como ausentes" in warning for warning in warnings))

    def test_ignore_room_preserves_date_rules_and_hides_room_columns_on_export(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            sistema = base / "sistema.xlsx"
            hotel = base / "hotel.xlsx"
            output = base / "resultado.xlsx"
            _write_rows(
                sistema,
                [["quarto", "nome", "check in", "check out"], ["101", "Maria Silva", "01/01/2026", "03/01/2026"]],
            )
            _write_rows(
                hotel,
                [["quarto", "nome", "check in", "check out"], ["999", "Maria Silva", "01/01/2026", "03/01/2026"]],
            )

            results, _, kpis = processar_arquivos(str(sistema), str(hotel), ignorar_quarto=True)
            write_bate_excel(results, output)
            wb = load_workbook(output, data_only=True)
            headers = [cell.value for cell in wb["RESULTADO COMPLETO"][2]]
            wb.close()

        self.assertEqual(results[0]["s_quarto"], Status.IGNORADO)
        self.assertEqual(results[0]["s_geral"], Status.OK)
        self.assertEqual(kpis["ok"], 1)
        self.assertNotIn("QUARTO (PLAN. 1)", headers)
        self.assertNotIn("QUARTO (PLAN. 2)", headers)
        self.assertNotIn("STATUS QUARTO", headers)

    def test_match_nomes_preserves_unmatched_and_empty_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            reference = base / "reference.xlsx"
            target = base / "target.xlsx"
            _write_rows(reference, [["nome"], ["Maria Silva"]])
            _write_rows(target, [["nome"], ["Pessoa Sem Par"], [None]])

            result = executar_match(str(reference), str(target), 95)

        self.assertTrue(result["ok"], result.get("erro"))
        self.assertEqual(result["nomes_finais"], ["Pessoa Sem Par", ""])
        self.assertEqual(result["statuses"], ["Nao encontrado", "NOME VAZIO"])
        self.assertEqual(result["kpis"], {"total": 2, "match": 0, "nomatch": 1, "empty": 1})


if __name__ == "__main__":
    unittest.main()
