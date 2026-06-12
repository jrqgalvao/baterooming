import re
import unicodedata
from functools import lru_cache

from rapidfuzz import fuzz


MATCH_FOUND = "Match encontrado"
NOT_FOUND = "Nao encontrado"
EMPTY_NAME = "NOME VAZIO"

_RE_NON_ALNUM_SPACE = re.compile(r"[^a-z0-9\s]")
_RE_SPACES = re.compile(r"\s+")
_STOPWORDS = {"de", "da", "do", "das", "dos", "e", "van", "von", "del", "di", "la", "le"}


@lru_cache(maxsize=8192)
def _normalizar_texto(nome: str) -> str:
    nome = unicodedata.normalize("NFD", nome)
    nome = "".join(c for c in nome if unicodedata.category(c) != "Mn")
    nome = nome.lower()
    nome = _RE_NON_ALNUM_SPACE.sub(" ", nome)
    nome = _RE_SPACES.sub(" ", nome).strip()
    return nome


def normalizar(nome: str) -> str:
    """Remove acentos, lowercase, pontuacao e espacos duplicados."""
    if not isinstance(nome, str) or nome.strip() == "":
        return ""
    return _normalizar_texto(nome)


def tokens_significativos(nome_norm: str) -> list[str]:
    return [t for t in nome_norm.split() if t not in _STOPWORDS and len(t) > 1]


def score_par(
    query: str,
    candidate: str,
    tq: "set | None" = None,
    tc: "set | None" = None,
) -> float:
    if not query or not candidate:
        return 0.0

    s_set = fuzz.token_set_ratio(query, candidate)
    s_sort = fuzz.token_sort_ratio(query, candidate)
    s_part = fuzz.partial_ratio(query, candidate)
    s_full = fuzz.ratio(query, candidate)

    score = 0.40 * s_set + 0.30 * s_sort + 0.20 * s_part + 0.10 * s_full

    if tq is None:
        tq = set(tokens_significativos(query))
    if tc is None:
        tc = set(tokens_significativos(candidate))
    if tq and tc:
        menor, maior = (tq, tc) if len(tq) <= len(tc) else (tc, tq)
        if menor and menor.issubset(maior):
            score = min(100.0, score + 5.0)

    return score


def _calcular_atribuicoes(
    nomes_b_normalizados: list[str],
    lista_a_normalizada: list[str],
    lista_a_tokens: list[set[str]],
    threshold: float,
    track_best_scores: bool = True,
) -> tuple[list[int], "list[float] | None"]:
    assignment_indices = [-1 for _ in nomes_b_normalizados]
    best_scores = [0.0 for _ in nomes_b_normalizados] if track_best_scores else None
    propostas: list[tuple[float, int, int]] = []
    linhas_b_atribuidas: set[int] = set()
    linhas_a_usadas: set[int] = set()

    exact_lookup: dict[str, list[int]] = {}
    for a_idx, nome_a_norm in enumerate(lista_a_normalizada):
        exact_lookup.setdefault(nome_a_norm, []).append(a_idx)

    for b_idx, nome_b_norm in enumerate(nomes_b_normalizados):
        if not nome_b_norm:
            continue
        for a_idx in exact_lookup.get(nome_b_norm, []):
            if a_idx in linhas_a_usadas:
                continue
            assignment_indices[b_idx] = a_idx
            if best_scores is not None:
                best_scores[b_idx] = 100.0
            linhas_b_atribuidas.add(b_idx)
            linhas_a_usadas.add(a_idx)
            break

    for b_idx, nome_b_norm in enumerate(nomes_b_normalizados):
        if not nome_b_norm or b_idx in linhas_b_atribuidas:
            continue

        tq = set(tokens_significativos(nome_b_norm))
        best_score = 0.0
        for a_idx, nome_a_norm in enumerate(lista_a_normalizada):
            if a_idx in linhas_a_usadas and not track_best_scores:
                continue
            score = score_par(nome_b_norm, nome_a_norm, tq, lista_a_tokens[a_idx])
            if best_scores is not None:
                best_score = max(best_score, score)
            if a_idx not in linhas_a_usadas and score >= threshold:
                propostas.append((score, b_idx, a_idx))
        if best_scores is not None:
            best_scores[b_idx] = round(best_score, 1)

    propostas.sort(key=lambda item: (-item[0], item[1], item[2]))

    for score, b_idx, a_idx in propostas:
        if b_idx in linhas_b_atribuidas or a_idx in linhas_a_usadas:
            continue
        assignment_indices[b_idx] = a_idx
        if best_scores is not None:
            best_scores[b_idx] = round(score, 1)
        linhas_b_atribuidas.add(b_idx)
        linhas_a_usadas.add(a_idx)

    return assignment_indices, best_scores


def atribuir_indices_matches(
    nomes_b_normalizados: list[str],
    lista_a_normalizada: list[str],
    lista_a_tokens: list[set[str]],
    threshold: float,
) -> list[int]:
    assignment_indices, _ = _calcular_atribuicoes(
        nomes_b_normalizados,
        lista_a_normalizada,
        lista_a_tokens,
        threshold,
        track_best_scores=False,
    )
    return assignment_indices


def atribuir_melhores_matches(
    nomes_b_originais: list[str],
    nomes_b_normalizados: list[str],
    lista_a_original: list[str],
    lista_a_normalizada: list[str],
    lista_a_tokens: list[set[str]],
    threshold: float,
):
    nomes_finais = list(nomes_b_originais)
    statuses = [EMPTY_NAME if not nome_norm else NOT_FOUND for nome_norm in nomes_b_normalizados]
    assignment_indices, scores = _calcular_atribuicoes(
        nomes_b_normalizados,
        lista_a_normalizada,
        lista_a_tokens,
        threshold,
    )
    assert scores is not None
    for b_idx, a_idx in enumerate(assignment_indices):
        if a_idx < 0:
            continue
        nomes_finais[b_idx] = lista_a_original[a_idx]
        statuses[b_idx] = MATCH_FOUND

    return nomes_finais, statuses, scores
