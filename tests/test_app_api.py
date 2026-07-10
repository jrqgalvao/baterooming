import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import app
from core.matching import ProcessingCancelled


class AppAPIRegressionTests(unittest.TestCase):
    def setUp(self):
        self.api = app.AppAPI()
        self.api._br_path1 = "br-1.xlsx"
        self.api._br_path2 = "br-2.xlsx"
        self.api._br_results = [{"nome": "Bate"}]
        self.api._mn_path1 = "mn-1.xlsx"
        self.api._mn_path2 = "mn-2.xlsx"
        self.api._mn_nomes_finais = ["Match"]
        self.api._mn_statuses = ["Match encontrado"]
        self.api._mn_scores = [100.0]
        self.api._mn_template_path = "mn-2.xlsx"
        self.api._mn_name_rows = [2]

    def test_limpar_bate_rooming_preserves_match_state(self):
        result = self.api.limpar()

        self.assertTrue(result["ok"])
        self.assertEqual(self.api._br_path1, "")
        self.assertEqual(self.api._br_path2, "")
        self.assertEqual(self.api._br_results, [])
        self.assertEqual(self.api._mn_path1, "mn-1.xlsx")
        self.assertEqual(self.api._mn_path2, "mn-2.xlsx")
        self.assertEqual(self.api._mn_nomes_finais, ["Match"])

    def test_limpar_match_preserves_bate_rooming_state(self):
        result = self.api.mn_limpar()

        self.assertTrue(result["ok"])
        self.assertEqual(self.api._mn_path1, "")
        self.assertEqual(self.api._mn_path2, "")
        self.assertEqual(self.api._mn_nomes_finais, [])
        self.assertEqual(self.api._br_path1, "br-1.xlsx")
        self.assertEqual(self.api._br_path2, "br-2.xlsx")
        self.assertEqual(self.api._br_results, [{"nome": "Bate"}])

    def test_file_slot_apis_reject_invalid_numbers_without_mutating_state(self):
        window = Mock()
        self.api._window = window

        responses = [
            self.api.selecionar_arquivo(3),
            self.api.limpar_arquivo(3),
            self.api.mn_selecionar_arquivo(3),
            self.api.mn_limpar_arquivo(3),
        ]

        self.assertTrue(all(response["ok"] is False for response in responses))
        self.assertTrue(all("1 ou 2" in response["erro"] for response in responses))
        self.assertEqual(self.api._br_path2, "br-2.xlsx")
        self.assertEqual(self.api._mn_path2, "mn-2.xlsx")
        window.create_file_dialog.assert_not_called()

    def test_file_slot_rejects_boolean_float_and_out_of_range_values(self):
        for invalid in (True, False, 1.0, 2.0, 0, "3", None):
            with self.subTest(invalid=invalid):
                self.assertIsNone(self.api._file_slot(invalid))

        self.assertEqual(self.api._file_slot(1), 1)
        self.assertEqual(self.api._file_slot(2), 2)
        self.assertEqual(self.api._file_slot("1"), 1)
        self.assertEqual(self.api._file_slot("2"), 2)

    def test_cancel_requests_only_signal_active_operations(self):
        self.assertFalse(self.api.cancelar_bate()["ok"])
        self.assertFalse(self.api.mn_cancelar()["ok"])

        self.api._br_running = True
        self.assertTrue(self.api.cancelar_bate()["ok"])
        self.assertTrue(self.api._br_cancel_event.is_set())

        self.api._mn_exporting = True
        self.assertTrue(self.api.mn_cancelar()["ok"])
        self.assertTrue(self.api._mn_cancel_event.is_set())

    def test_bate_export_writes_temporary_file_before_publishing(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            target = base / "resultado.xlsx"
            self.api._br_path1 = str(base / "origem_1.xlsx")
            self.api._br_path2 = str(base / "origem_2.xlsx")
            self.api._window = Mock()
            self.api._window.create_file_dialog.return_value = (str(target),)
            written_paths = []

            def fake_write_excel(data, path, **kwargs):
                output = Path(path)
                written_paths.append(output)
                self.assertNotEqual(output.resolve(), target.resolve())
                output.write_text("novo", encoding="utf-8")

            with patch.object(
                app,
                "_load_bate_rooming_services",
                return_value=(Mock(), fake_write_excel),
            ):
                result = self.api.exportar()

            self.assertTrue(result["ok"])
            self.assertEqual(target.read_text(encoding="utf-8"), "novo")
            self.assertEqual(len(written_paths), 1)
            self.assertFalse(written_paths[0].exists())

    def test_bate_export_cancellation_keeps_existing_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            target = base / "resultado.xlsx"
            target.write_text("anterior", encoding="utf-8")
            self.api._br_path1 = str(base / "origem_1.xlsx")
            self.api._br_path2 = str(base / "origem_2.xlsx")
            self.api._window = Mock()
            self.api._window.create_file_dialog.return_value = (str(target),)

            def cancel_write_excel(data, path, **kwargs):
                Path(path).write_text("parcial", encoding="utf-8")
                self.api._br_cancel_event.set()
                raise ProcessingCancelled()

            with patch.object(
                app,
                "_load_bate_rooming_services",
                return_value=(Mock(), cancel_write_excel),
            ):
                result = self.api.exportar()

            self.assertFalse(result["ok"])
            self.assertTrue(result["cancelado"])
            self.assertEqual(target.read_text(encoding="utf-8"), "anterior")
            self.assertEqual(list(base.glob("*.xlsx")), [target])

    def test_exports_refuse_to_replace_source_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "origem.xlsx"
            self.api._br_path1 = str(source)
            self.api._br_path2 = str(Path(tmp) / "origem_2.xlsx")
            self.api._mn_path1 = str(source)
            self.api._mn_path2 = str(Path(tmp) / "origem_match.xlsx")
            self.api._window = Mock()
            self.api._window.create_file_dialog.return_value = (str(source),)
            bate_writer = Mock()
            match_writer = Mock()

            with patch.object(
                app,
                "_load_bate_rooming_services",
                return_value=(Mock(), bate_writer),
            ):
                bate_result = self.api.exportar()
            with patch.object(
                app,
                "_load_match_nomes_services",
                return_value=(Mock(), match_writer),
            ):
                match_result = self.api.mn_exportar()

        self.assertFalse(bate_result["ok"])
        self.assertFalse(match_result["ok"])
        self.assertIn("não substituir", bate_result["erro"])
        self.assertIn("não substituir", match_result["erro"])
        bate_writer.assert_not_called()
        match_writer.assert_not_called()

    def test_main_uses_menu_profile_for_initial_window_size(self):
        window = Mock()
        with (
            patch.object(app, "_resolve") as resolve,
            patch.object(app, "_resolve_url", return_value="file:///menu_ui.html"),
            patch.object(app, "_resolve_page_size", return_value=(640, 520)) as resolve_size,
            patch.object(app, "_center_coords", return_value=(100, 100)),
            patch.object(app.webview, "create_window", return_value=window),
            patch.object(app.webview, "start"),
        ):
            app.main()

        resolve_size.assert_called_once_with("menu_ui.html")
        resolve.assert_any_call("menu_ui.html")


if __name__ == "__main__":
    unittest.main()
