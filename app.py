"""
app.py - {{PRODUCT_NAME}} v1.4.4
Entry point unificado: Menu inicial + Bate-Rooming + Match de Nomes.

Usa UMA janela pywebview que navega entre arquivos locais via load_url(),
evitando o NavigateToString/load_html que pode exibir frames brancos.

Estrutura de arquivos necessária:
    app.py
    core/
    ui/
    assets/

Dependências:
    pip install pywebview openpyxl rapidfuzz

Para gerar .exe:
    pyinstaller --noconfirm app.spec
"""

import os
import platform
import re
import subprocess
import sys
from datetime import date
from pathlib import Path
from threading import Lock

try:
    import webview
except ImportError:
    print("ERRO: pywebview não encontrado. Execute: pip install pywebview")
    input("Pressione Enter para sair...")
    sys.exit(1)

_bate_rooming_services = None
_match_nomes_services = None


def _load_bate_rooming_services():
    global _bate_rooming_services
    if _bate_rooming_services is None:
        from core.bate_rooming import processar_arquivos
        from core.bate_rooming_export import write_excel

        _bate_rooming_services = (processar_arquivos, write_excel)
    return _bate_rooming_services


def _load_match_nomes_services():
    global _match_nomes_services
    if _match_nomes_services is None:
        from core.match_nomes import executar_match
        from core.match_nomes_export import write_excel

        _match_nomes_services = (executar_match, write_excel)
    return _match_nomes_services


# ── Resolução de caminhos ──────────────────────────────────────
def _base() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).parent


def _resolve(filename: str) -> str:
    ui_files = {"menu_ui.html", "bate_rooming_ui.html", "match_nomes_ui.html"}
    asset_files = {"app_icon.ico", "logo_generic_color.png", "logo_generic_white.png"}
    if filename in ui_files:
        p = _base() / "ui" / filename
    elif filename in asset_files:
        p = _base() / "assets" / filename
    else:
        p = _base() / filename
    if not p.exists():
        raise FileNotFoundError(
            f"Interface não encontrada: {filename}\n"
            f"Certifique-se de que está na mesma pasta que app.py."
        )
    return str(p)


def _resolve_url(filename: str) -> str:
    return Path(_resolve(filename)).resolve().as_uri()


def _open_file(path: str) -> dict:
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return {"ok": True}
    except Exception as exc:
        return {"ok": False, "erro": _friendly_error(exc, "abrir arquivo")}


def _open_folder(path: str) -> dict:
    p = Path(path)
    folder = p if p.is_dir() else p.parent
    if not folder.exists():
        return {"ok": False, "erro": "Não encontramos a pasta desse arquivo. Salve o relatório novamente e tente abrir a pasta depois."}
    return _open_file(str(folder))


def _slug_filename_part(value: str) -> str:
    stem = Path(value).stem if value else "arquivo"
    slug = re.sub(r"[^A-Za-z0-9]+", "_", stem).strip("_").lower()
    return slug[:34] or "arquivo"


def _default_export_name(prefix: str, paths: list[str] | tuple[str, ...]) -> str:
    parts = [_slug_filename_part(path) for path in paths if path]
    joined = "_".join(parts[:2]) if parts else "resultado"
    return f"{prefix}_{joined}_{date.today().isoformat()}.xlsx"


def _match_export_name(planilha_2_path: str) -> str:
    stem = Path(planilha_2_path).stem.strip() if planilha_2_path else "Planilha 2"
    return f"{stem or 'Planilha 2'}_corrigido.xlsx"


def _friendly_error(exc: Exception, context: str = "arquivo") -> str:
    msg = str(exc).strip()
    lower = msg.lower()
    generic = (
        f"Não conseguimos continuar no {context}. "
        "Tente novamente. Se acontecer de novo, feche o app, abra novamente e repita a última ação."
    )
    technical_markers = (
        "traceback",
        "typeerror",
        "valueerror",
        "appapi.",
        "positional argument",
        "takes ",
        "were given",
        "nonetype",
        "object has no attribute",
    )
    if not msg:
        return generic
    if any(marker in lower for marker in technical_markers):
        return generic
    if isinstance(exc, PermissionError) or "permission denied" in lower:
        return "Não conseguimos salvar porque o arquivo está aberto em outro programa. Feche o Excel e tente salvar de novo."
    if "not a zip file" in lower or "file is not a zip" in lower:
        return (
            "O arquivo selecionado não parece ser uma planilha Excel válida. "
            "Confirme se escolheu um arquivo .xlsx/.xlsm correto e tente novamente."
        )
    if "no sheet" in lower or "worksheet" in lower:
        return "A planilha parece estar vazia ou sem abas válidas. Verifique o arquivo e tente novamente."
    if "coluna" in lower or "column" in lower:
        return (
            f"{msg}\n\nIsso normalmente acontece quando o arquivo errado foi selecionado "
            f"ou quando o cabeçalho da planilha não está no formato esperado para {context}."
        )
    if "empty" in lower or "vazia" in lower or "vazio" in lower:
        return "A planilha não tem dados para processar. Verifique se o arquivo correto foi selecionado."
    return msg


def _file_summary(path: str, allowed_exts: tuple[str, ...]) -> dict:
    p = Path(path)
    ext = p.suffix.lower()
    try:
        size = p.stat().st_size
    except OSError:
        size = 0

    if size >= 1024 * 1024:
        size_label = f"{size / (1024 * 1024):.1f} MB"
    elif size >= 1024:
        size_label = f"{size / 1024:.0f} KB"
    else:
        size_label = f"{size} B"

    is_expected = ext in allowed_exts
    return {
        "ext": ext or "sem extensao",
        "size": size,
        "size_label": size_label,
        "expected": is_expected,
        "detail": f"{(ext or 'sem extensao').upper()} · {size_label} · "
                  f"{'formato esperado' if is_expected else 'verifique o formato'}",
    }



def _get_primary_screen():
    """Retorna (x, y, width, height) do monitor principal (origem 0,0)."""
    try:
        screens = webview.screens
        primary = next(
            (s for s in screens
             if int(getattr(s, "x", -1)) == 0 and int(getattr(s, "y", -1)) == 0),
            screens[0]
        )
        return (
            int(getattr(primary, "x", 0)),
            int(getattr(primary, "y", 0)),
            int(primary.width),
            int(primary.height),
        )
    except Exception:
        return (0, 0, 1920, 1080)


def _screen_for_window(window, fallback=None):
    """Retorna o monitor onde a janela esta, preservando setups com 2 monitores."""
    fallback = fallback or _get_primary_screen()
    try:
        wx = int(getattr(window, "x", fallback[0]))
        wy = int(getattr(window, "y", fallback[1]))
        ww = int(getattr(window, "width", 0) or 0)
        wh = int(getattr(window, "height", 0) or 0)
        cx = wx + max(ww, 1) // 2
        cy = wy + max(wh, 1) // 2

        for screen in webview.screens:
            sx = int(getattr(screen, "x", 0))
            sy = int(getattr(screen, "y", 0))
            sw = int(screen.width)
            sh = int(screen.height)
            if sx <= cx < sx + sw and sy <= cy < sy + sh:
                return sx, sy, sw, sh
    except Exception:
        pass
    return fallback


def _center_coords(win_w: int, win_h: int, window=None, screen=None):
    """Calcula (x, y) centralizando no monitor atual da janela."""
    ox, oy, sw, sh = screen or (_screen_for_window(window) if window else _get_primary_screen())
    return ox + max(0, (sw - win_w) // 2), oy + max(0, (sh - win_h) // 2)


_SCREEN_MARGIN = 80
_WINDOW_PROFILES = {
    "menu_ui.html": {
        "width_pct": 0.42,
        "height_pct": 0.56,
        "min": (520, 420),
        "max": (680, 560),
    },
    "match_nomes_ui.html": {
        "width_pct": 0.62,
        "height_pct": 0.74,
        "min": (680, 540),
        "max": (980, 760),
    },
    "bate_rooming_ui.html": {
        "width_pct": 0.86,
        "height_pct": 0.84,
        "min": (960, 640),
        "max": (1440, 900),
    },
}


def _bounded_dimension(ideal: int, min_value: int, max_value: int) -> int:
    if max_value < min_value:
        return max_value
    return max(min_value, min(ideal, max_value))


def _resolve_page_size(page: str, window=None, screen=None) -> tuple[int, int]:
    """Calcula tamanho responsivo por pagina usando o monitor ativo."""
    page = Path(str(page)).name
    profile = _WINDOW_PROFILES.get(page, _WINDOW_PROFILES["menu_ui.html"])
    _, _, sw, sh = screen or (_screen_for_window(window) if window else _get_primary_screen())

    max_w = min(profile["max"][0], max(480, sw - _SCREEN_MARGIN))
    max_h = min(profile["max"][1], max(320, sh - _SCREEN_MARGIN))
    ideal_w = round(sw * profile["width_pct"])
    ideal_h = round(sh * profile["height_pct"])

    return (
        _bounded_dimension(ideal_w, profile["min"][0], max_w),
        _bounded_dimension(ideal_h, profile["min"][1], max_h),
    )


# ─────────────────────────────────────────────────────────────
# API UNIFICADA — exposta ao JS de todas as páginas
# ─────────────────────────────────────────────────────────────
class AppAPI:
    """
    Única API compartilhada por todas as telas.
    A janela é única; a troca de ferramenta muda apenas o HTML carregado.
    """

    def __init__(self):
        self._window = None          # definido após create_window
        self._nav_lock = Lock()
        self._current_page = "menu_ui.html"

        # ── estado Bate-Rooming ──
        self._br_results: list = []
        self._br_path1: str   = ""
        self._br_path2: str   = ""

        # ── estado Match de Nomes ──
        self._mn_path1: str       = ""
        self._mn_path2: str       = ""
        self._mn_nomes_finais     = []
        self._mn_statuses         = []
        self._mn_scores           = []
        self._mn_template_path     = ""
        self._mn_name_rows         = []
        self._mn_name_column       = 2

    # ── Navegação ─────────────────────────────────────────────
    def _wait_for_loaded(self) -> None:
        loaded = getattr(getattr(self._window, "events", None), "loaded", None)
        wait = getattr(loaded, "wait", None)
        if callable(wait):
            wait(timeout=3)

    def _navigate(self, filename: str) -> dict:
        """Navega por URL local para evitar o branco do NavigateToString."""
        with self._nav_lock:
            if filename == self._current_page:
                return {"ok": True, "page": filename, "skipped": True}

            previous_page = self._current_page
            try:
                self._window.load_url(_resolve_url(filename))
                self._wait_for_loaded()
                self._current_page = filename
                return {"ok": True, "page": filename}
            except Exception as exc:
                self._current_page = previous_page
                return {"ok": False, "erro": _friendly_error(exc, "abrir tela")}

    def abrir_ferramenta(self, tool: str) -> dict:
        """Chamado pelo menu: navega para a ferramenta."""
        if tool == "bate_rooming":
            return self._navigate("bate_rooming_ui.html")
        if tool == "match_nomes":
            return self._navigate("match_nomes_ui.html")
        return {"ok": False, "erro": "Não encontramos essa ferramenta. Volte ao menu e escolha uma opção disponível."}

    def voltar_menu(self) -> dict:
        """Volta ao menu inicial."""
        return self._navigate("menu_ui.html")

    def abrir_arquivo(self, path: str) -> dict:
        return _open_file(path)

    def abrir_pasta(self, path: str) -> dict:
        return _open_folder(path)

    @staticmethod
    def _serialize_bate_result(r: dict) -> dict:
        return {
            "nome":     r["nome"],     "fonte":    r["fonte"],
            "no_match": r["no_match"], "is_dup":   r.get("is_dup", False),
            "room_check_ignored": r.get("room_check_ignored", False),
            "q_sys":    r["q_sys"],    "q_hotel":  r["q_hotel"],
            "ci_sys":   r["ci_sys"],   "ci_hotel": r["ci_hotel"],
            "co_sys":   r["co_sys"],   "co_hotel": r["co_hotel"],
            "s_quarto": str(r["s_quarto"]),
            "s_ci":     str(r["s_ci"]), "s_co":    str(r["s_co"]),
        }

    def _clear_bate_results(self) -> None:
        self._br_results = []

    def _clear_bate_state(self) -> None:
        self._br_path1 = ""
        self._br_path2 = ""
        self._clear_bate_results()

    def _clear_match_results(self) -> None:
        self._mn_nomes_finais = []
        self._mn_statuses     = []
        self._mn_scores       = []
        self._mn_template_path = ""
        self._mn_name_rows     = []
        self._mn_name_column   = 2

    def _clear_match_state(self) -> None:
        self._mn_path1 = ""
        self._mn_path2 = ""
        self._clear_match_results()

    @staticmethod
    def _file_slot(numero) -> int | None:
        if isinstance(numero, bool):
            return None
        if isinstance(numero, int) and numero in (1, 2):
            return numero
        if isinstance(numero, str) and numero in ("1", "2"):
            return int(numero)
        return None

    @staticmethod
    def _invalid_file_slot() -> dict:
        return {"ok": False, "erro": "O número do arquivo precisa ser 1 ou 2."}

    def _emit_progress(self, percent: int, message: str, status_type: str = "spin") -> None:
        if not self._window:
            return
        script = (
            "if (window.onPythonProgress) "
            f"window.onPythonProgress({int(percent)}, {repr(str(message))}, {repr(str(status_type))});"
        )
        try:
            self._window.evaluate_js(script)
        except Exception:
            pass

    @staticmethod
    def _single_dialog_path(result) -> str:
        if not result:
            return ""
        if isinstance(result, (str, Path)):
            return str(result)
        return str(result[0]) if result else ""

    @staticmethod
    def _dialog_directory(*paths: str) -> str:
        for raw_path in paths:
            if not raw_path:
                continue
            path = Path(raw_path).expanduser()
            directory = path if path.is_dir() else path.parent
            if directory.exists():
                return str(directory)

        downloads = Path.home() / "Downloads"
        return str(downloads if downloads.exists() else Path.home())

    # ─────────────────────────────────────────────────────────
    # BATE-ROOMING
    # ─────────────────────────────────────────────────────────
    def selecionar_arquivo(self, numero: int) -> dict:
        numero = self._file_slot(numero)
        if numero is None:
            return self._invalid_file_slot()
        try:
            old_path = self._br_path1 if numero == 1 else self._br_path2
            result = self._window.create_file_dialog(
                webview.OPEN_DIALOG,
                directory=self._dialog_directory(old_path),
                allow_multiple=False,
                file_types=("Excel (*.xlsx;*.xlsm;*.xls)", "Todos os arquivos (*.*)")
            )
            if not result:
                return {"ok": False, "erro": "Nenhum arquivo selecionado."}
            path = self._single_dialog_path(result)
            changed = path != old_path
            if numero == 1:
                self._br_path1 = path
            else:
                self._br_path2 = path
            if changed:
                self._clear_bate_results()
            return {"ok": True, "path": path, "nome": Path(path).name, "changed": changed,
                    "file_info": _file_summary(path, (".xlsx", ".xlsm", ".xls"))}
        except Exception as exc:
            return {"ok": False, "erro": _friendly_error(exc, "selecionar arquivo")}

    def executar(self, options: dict | None = None) -> dict:
        if not self._br_path1 or not self._br_path2:
            return {"ok": False, "erro": "Selecione as duas planilhas antes de executar."}
        for p in (self._br_path1, self._br_path2):
            if not Path(p).exists():
                return {"ok": False, "erro": "Não encontramos um dos arquivos selecionados. Escolha as planilhas novamente e tente executar."}
        processar_arquivos, _ = _load_bate_rooming_services()
        try:
            ignorar_quarto = bool((options or {}).get("ignorar_quarto"))
            self._emit_progress(30, "Validando planilhas e aplicando match inteligente...")
            results, warnings, kpis = processar_arquivos(
                self._br_path1,
                self._br_path2,
                ignorar_quarto=ignorar_quarto,
            )
            self._emit_progress(85, "Preparando resultado para a tela...")
            self._br_results = results
            serialized = [self._serialize_bate_result(r) for r in results]
            return {"ok": True, "results": serialized, "kpis": kpis, "warnings": warnings}
        except ValueError as exc:
            return {"ok": False, "erro": _friendly_error(exc, "Bate-Rooming")}
        except Exception as exc:
            return {"ok": False, "erro": _friendly_error(exc, "Bate-Rooming")}

    def exportar(self, apenas_visiveis: bool = False, indices_visiveis=None) -> dict:
        if not self._br_results:
            return {"ok": False, "erro": "Nenhum resultado. Execute o Bate-Rooming primeiro."}
        data = self._br_results
        if apenas_visiveis and indices_visiveis is not None:
            try:
                indices = [int(i) for i in indices_visiveis]
                data = [self._br_results[i] for i in indices if 0 <= i < len(self._br_results)]
            except Exception:
                data = self._br_results
        _, br_write_excel = _load_bate_rooming_services()
        try:
            result = self._window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=_default_export_name("bate_rooming", (self._br_path1, self._br_path2)),
                file_types=("Excel (*.xlsx)",)
            )
            if not result:
                return {"ok": False, "erro": "Exportação cancelada."}
            path = result if isinstance(result, str) else result[0]
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            br_write_excel(data, Path(path))
            return {"ok": True, "path": path, "nome": Path(path).name}
        except PermissionError as exc:
            return {"ok": False, "erro": _friendly_error(exc, "Bate-Rooming")}
        except Exception as exc:
            return {"ok": False, "erro": _friendly_error(exc, "Bate-Rooming")}

    def limpar(self) -> dict:
        self._clear_bate_state()
        return {"ok": True}

    def limpar_arquivo(self, numero: int) -> dict:
        numero = self._file_slot(numero)
        if numero is None:
            return self._invalid_file_slot()
        if numero == 1:
            self._br_path1 = ""
        else:
            self._br_path2 = ""
        self._clear_bate_results()
        return {"ok": True}

    # ─────────────────────────────────────────────────────────
    # MATCH DE NOMES
    # ─────────────────────────────────────────────────────────
    def mn_selecionar_arquivo(self, numero: int) -> dict:
        numero = self._file_slot(numero)
        if numero is None:
            return self._invalid_file_slot()
        try:
            old_path = self._mn_path1 if numero == 1 else self._mn_path2
            result = self._window.create_file_dialog(
                webview.OPEN_DIALOG,
                directory=self._dialog_directory(old_path, self._mn_path1, self._mn_path2),
                allow_multiple=False,
                file_types=("Excel (*.xlsx)", "Todos os arquivos (*.*)")
            )
            if not result:
                return {"ok": False, "erro": "Nenhum arquivo selecionado."}
            path = self._single_dialog_path(result)
            changed = path != old_path
            if numero == 1:
                self._mn_path1 = path
            else:
                self._mn_path2 = path
            if changed:
                self._clear_match_results()
            return {"ok": True, "path": path, "nome": Path(path).name, "changed": changed,
                    "file_info": _file_summary(path, (".xlsx",))}
        except Exception as exc:
            return {"ok": False, "erro": _friendly_error(exc, "selecionar arquivo")}

    def mn_executar(self, threshold: int = 65) -> dict:
        if not self._mn_path1 or not self._mn_path2:
            return {"ok": False, "erro": "Selecione as duas listas antes de executar."}
        for p in (self._mn_path1, self._mn_path2):
            if not Path(p).exists():
                return {"ok": False, "erro": "Não encontramos um dos arquivos selecionados. Escolha as planilhas novamente e tente executar."}
        try:
            threshold = max(50, min(95, int(threshold)))
        except (TypeError, ValueError):
            return {"ok": False, "erro": "A sensibilidade precisa ser um número entre 50 e 95. Ajuste o controle e tente novamente."}
        executar_match, _ = _load_match_nomes_services()
        try:
            self._emit_progress(35, "Lendo listas e comparando nomes...")
            resultado = executar_match(self._mn_path1, self._mn_path2, threshold)
            self._emit_progress(85, "Preparando resumo do match...")
        except Exception as exc:
            return {"ok": False, "erro": _friendly_error(exc, "Match de Nomes")}
        if not resultado["ok"]:
            return {"ok": False, "erro": _friendly_error(ValueError(resultado.get("erro", "")), "Match de Nomes")}
        if resultado["kpis"].get("total", 0) == 0:
            return {"ok": False, "erro": "A Planilha 2 não tem nomes para processar. Verifique se selecionou o arquivo correto."}
        return self._store_match_result(resultado)

    def _store_match_result(self, resultado: dict) -> dict:
        self._mn_nomes_finais = resultado["nomes_finais"]
        self._mn_statuses     = resultado["statuses"]
        self._mn_scores       = resultado.get("scores", [])
        self._mn_template_path = resultado.get("template_path", self._mn_path2)
        self._mn_name_rows     = resultado.get("name_rows", [])
        self._mn_name_column   = resultado.get("name_column", 2)
        return {"ok": True, "kpis": resultado["kpis"]}

    def mn_exportar(self) -> dict:
        if not self._mn_nomes_finais:
            return {"ok": False, "erro": "Nenhum resultado. Execute o Match primeiro."}
        _, mn_write_excel = _load_match_nomes_services()
        try:
            result = self._window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=_match_export_name(self._mn_path2),
                file_types=("Excel (*.xlsx)",)
            )
            if not result:
                return {"ok": False, "erro": "Exportação cancelada."}
            path = result if isinstance(result, str) else result[0]
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            mn_write_excel(
                self._mn_nomes_finais,
                self._mn_statuses,
                Path(path),
                self._mn_scores,
                template_path=self._mn_template_path,
                name_rows=self._mn_name_rows,
                name_column=self._mn_name_column,
            )
            return {"ok": True, "path": path, "nome": Path(path).name}
        except PermissionError as exc:
            return {"ok": False, "erro": _friendly_error(exc, "Match de Nomes")}
        except Exception as exc:
            return {"ok": False, "erro": _friendly_error(exc, "Match de Nomes")}

    def mn_limpar(self) -> dict:
        self._clear_match_state()
        return {"ok": True}

    def mn_limpar_arquivo(self, numero: int) -> dict:
        numero = self._file_slot(numero)
        if numero is None:
            return self._invalid_file_slot()
        if numero == 1:
            self._mn_path1 = ""
        else:
            self._mn_path2 = ""
        self._clear_match_results()
        return {"ok": True}


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────
def main():
    # Valida presença dos HTMLs antes de iniciar
    for fname in ("menu_ui.html", "bate_rooming_ui.html", "match_nomes_ui.html"):
        try:
            _resolve(fname)
        except FileNotFoundError as exc:
            print(f"ERRO: {exc}")
            input("Pressione Enter para sair...")
            sys.exit(1)

    api    = AppAPI()
    _init_w, _init_h = _resolve_page_size("menu_ui.html")
    _init_x, _init_y = _center_coords(_init_w, _init_h)
    window = webview.create_window(
        title      = "{{PRODUCT_NAME}}",
        url        = _resolve_url("menu_ui.html"),
        js_api     = api,
        width      = _init_w,
        height     = _init_h,
        x          = _init_x,
        y          = _init_y,
        min_size   = (480, 320),
        resizable  = True,
        frameless  = False,
        easy_drag  = False,
        text_select= False,
        background_color="#F6F7F6",
    )
    api._window = window
    webview.start(debug=False)


if __name__ == "__main__":
    main()

