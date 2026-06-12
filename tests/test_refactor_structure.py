import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class RefactorStructureTests(unittest.TestCase):
    def test_runtime_and_development_dependencies_are_separated(self):
        runtime = (ROOT / "requirements.txt").read_text(encoding="utf-8").splitlines()
        development = (ROOT / "requirements-dev.txt").read_text(encoding="utf-8").splitlines()

        self.assertNotIn("pandas", runtime)
        self.assertNotIn("pytest", runtime)
        self.assertNotIn("pyinstaller", runtime)
        self.assertIn("-r requirements.txt", development)
        self.assertIn("pytest", development)
        self.assertIn("pyinstaller", development)

    def test_generated_artifacts_are_ignored(self):
        gitignore = (ROOT / ".gitignore").read_text(encoding="utf-8")

        for pattern in ("__pycache__/", "*.py[cod]", ".pytest_cache/", ".ruff_cache/", "build/", "dist/"):
            with self.subTest(pattern=pattern):
                self.assertIn(pattern, gitignore)

    def test_project_uses_core_ui_assets_layout(self):
        expected_paths = [
            "app.py",
            "core/__init__.py",
            "core/matching.py",
            "core/bate_rooming.py",
            "core/bate_rooming_export.py",
            "core/match_nomes.py",
            "core/match_nomes_export.py",
            "ui/menu_ui.html",
            "ui/bate_rooming_ui.html",
            "ui/match_nomes_ui.html",
            "assets/app_icon.ico",
            "assets/logo_generic_color.png",
            "assets/logo_generic_white.png",
        ]
        for relative_path in expected_paths:
            with self.subTest(path=relative_path):
                self.assertTrue((ROOT / relative_path).exists())

    def test_root_no_longer_contains_moved_ui_or_asset_files(self):
        moved_root_files = [
            "menu_ui.html",
            "bate_rooming_ui.html",
            "match_nomes_ui.html",
            "app_icon.ico",
            "logo_generic_color.png",
            "logo_generic_white.png",
            "bate_rooming_core.py",
            "match_nomes_core.py",
        ]
        for relative_path in moved_root_files:
            with self.subTest(path=relative_path):
                self.assertFalse((ROOT / relative_path).exists())

    def test_pyinstaller_spec_points_to_new_layout(self):
        spec = (ROOT / "app.spec").read_text(encoding="utf-8")
        self.assertIn("ui/menu_ui.html", spec)
        self.assertIn("ui/bate_rooming_ui.html", spec)
        self.assertIn("ui/match_nomes_ui.html", spec)
        self.assertIn("assets/app_icon.ico", spec)
        self.assertIn("assets/logo_generic_color.png", spec)
        self.assertIn("assets/logo_generic_white.png", spec)
        self.assertNotIn("('menu_ui.html', '.')", spec)
        self.assertNotIn("('logo_generic_color.png', '.')", spec)
        self.assertIn("excludes=['pytest', 'numpy']", spec)
        self.assertIn("version='app_version_info.txt'", spec)

    def test_program_version_is_1_4_4(self):
        app_source = (ROOT / "app.py").read_text(encoding="utf-8")
        menu = (ROOT / "ui" / "menu_ui.html").read_text(encoding="utf-8")
        version_info = (ROOT / "app_version_info.txt").read_text(encoding="utf-8")

        self.assertIn("{{PRODUCT_NAME}} v1.4.4", app_source)
        self.assertIn(">v1.4.4</span>", menu)
        self.assertIn("Abrir changelog da versão 1.4.4", menu)
        self.assertIn("filevers=(1, 4, 4, 0)", version_info)
        self.assertIn("prodvers=(1, 4, 4, 0)", version_info)
        self.assertIn("StringStruct('ProductVersion', '1.4.4')", version_info)

    def test_match_nomes_has_no_pandas_or_residual_ui_helpers(self):
        core_source = (ROOT / "core" / "match_nomes.py").read_text(encoding="utf-8")
        html = (ROOT / "ui" / "match_nomes_ui.html").read_text(encoding="utf-8")

        self.assertNotIn("import pandas", core_source)
        self.assertNotIn("pd.", core_source)
        self.assertNotIn("function updateResultLink", html)
        self.assertNotIn("function abrirArquivo", html)
        self.assertNotIn("function abrirPasta", html)

    def test_ui_pages_keep_required_visual_assets(self):
        expected_assets = {
            "menu_ui.html": ["../assets/logo_generic_color.png", "../assets/logo_generic_white.png"],
            "bate_rooming_ui.html": ["../assets/logo_generic_color.png", "../assets/logo_generic_white.png"],
            "match_nomes_ui.html": ["../assets/logo_generic_color.png", "../assets/logo_generic_white.png"],
        }
        for page, assets in expected_assets.items():
            html = (ROOT / "ui" / page).read_text(encoding="utf-8")
            for asset in assets:
                with self.subTest(page=page, asset=asset):
                    self.assertIn(asset, html)

    def test_menu_interactive_elements_support_keyboard_and_changelog_access(self):
        html = (ROOT / "ui" / "menu_ui.html").read_text(encoding="utf-8")

        self.assertIn('class="version"', html)
        self.assertIn('role="button"', html)
        self.assertIn('tabindex="0"', html)
        self.assertIn('onclick="openChangelog()"', html)
        self.assertIn("function openChangelog()", html)
        self.assertEqual(html.count('onkeydown="handleToolKey(event,'), 2)

    def test_primary_ui_actions_have_accessible_names(self):
        expected_labels = {
            "bate_rooming_ui.html": (
                'id="btn-run" aria-label="Executar Bate-Rooming"',
                'id="btn-export" aria-label="Exportar resultado do Bate-Rooming para Excel"',
            ),
            "match_nomes_ui.html": (
                'id="btn-run" aria-label="Executar Match de Nomes"',
                'id="btn-export" aria-label="Exportar resultado do Match de Nomes para Excel"',
            ),
        }
        for page, labels in expected_labels.items():
            html = (ROOT / "ui" / page).read_text(encoding="utf-8")
            for label in labels:
                with self.subTest(page=page, label=label):
                    self.assertIn(label, html)

    def test_ui_pages_are_loaded_as_file_urls(self):
        import app

        url = app._resolve_url("menu_ui.html")

        self.assertTrue(url.startswith("file:///"))
        self.assertIn("/ui/menu_ui.html", url)


if __name__ == "__main__":
    unittest.main()
