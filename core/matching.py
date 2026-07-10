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
_DETAIL_TOP_LIMIT = 5


class ProcessingCancelled(Exception):
    pass


def check_cancel(should_cancel=None) -> None:
    if should_cancel is not None and should_cancel():
        raise ProcessingCancelled()


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


def _identidade_nome(nome_norm: str) -> "tuple[str, str] | None":
    tokens = tokens_significativos(nome_norm)
    return (tokens[0], tokens[-1]) if len(tokens) >= 2 else None


def _identidades_compativeis(
    query_identity: "tuple[str, str] | None",
    candidate_identity: "tuple[str, str] | None",
    minimum: float,
) -> bool:
    if query_identity is None or candidate_identity is None:
        return False
    return (
        fuzz.ratio(query_identity[0], candidate_identity[0]) >= minimum
        and fuzz.ratio(query_identity[1], candidate_identity[1]) >= minimum
    )


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
    progress=None,
    should_cancel=None,
    min_cross_identity_score: float | None = None,
    match_details: list[dict] | None = None,
) -> tuple[list[int], "list[float] | None"]:
    assignment_indices = [-1 for _ in nomes_b_normalizados]
    best_scores = [0.0 for _ in nomes_b_normalizados] if track_best_scores else None
    propostas: list[tuple[float, int, int]] = []
    linhas_b_atribuidas: set[int] = set()
    linhas_a_usadas: set[int] = set()
    assigned_stage = ["SEM_MATCH" for _ in nomes_b_normalizados]
    assigned_scores = [None for _ in nomes_b_normalizados]
    assigned_by_a: dict[int, int] = {}
    conflicts: list[list[dict]] = [[] for _ in nomes_b_normalizados]
    candidate_considered = [0 for _ in nomes_b_normalizados]
    candidate_eligible = [0 for _ in nomes_b_normalizados]
    candidate_blocked = [0 for _ in nomes_b_normalizados]
    candidate_top: list[list[dict]] = [[] for _ in nomes_b_normalizados]
    if match_details is not None:
        match_details[:] = [{} for _ in nomes_b_normalizados]
    total = max(len(nomes_b_normalizados), 1)

    def report(fraction: float) -> None:
        if progress is not None:
            progress(min(1.0, fraction))

    def record_candidate(
        b_idx: int,
        a_idx: int,
        score: float,
        stage: str,
        *,
        eligible: bool = False,
        blocked: bool = False,
        reason: str = "",
    ) -> None:
        if match_details is None:
            return
        candidate_considered[b_idx] += 1
        if eligible:
            candidate_eligible[b_idx] += 1
        if blocked:
            candidate_blocked[b_idx] += 1
        item = {
            "a_idx": a_idx,
            "score": round(score, 1),
            "stage": stage,
            "reason": reason,
        }
        top = candidate_top[b_idx]
        for index, existing in enumerate(top):
            if existing["a_idx"] != a_idx:
                continue
            if existing["score"] >= item["score"]:
                return
            top.pop(index)
            break
        top.append(item)
        top.sort(key=lambda value: (-value["score"], value["a_idx"]))
        del top[_DETAIL_TOP_LIMIT:]

    def assign(b_idx: int, a_idx: int, score: float, stage: str) -> None:
        assignment_indices[b_idx] = a_idx
        linhas_b_atribuidas.add(b_idx)
        linhas_a_usadas.add(a_idx)
        assigned_stage[b_idx] = stage
        assigned_scores[b_idx] = round(score, 1)
        assigned_by_a[a_idx] = b_idx
        if best_scores is not None:
            best_scores[b_idx] = round(score, 1)

    exact_lookup: dict[str, list[int]] = {}
    for a_idx, nome_a_norm in enumerate(lista_a_normalizada):
        exact_lookup.setdefault(nome_a_norm, []).append(a_idx)

    for b_idx, nome_b_norm in enumerate(nomes_b_normalizados):
        if b_idx % 128 == 0:
            check_cancel(should_cancel)
            report(0.2 * b_idx / total)
        if not nome_b_norm:
            continue
        for a_idx in exact_lookup.get(nome_b_norm, []):
            if a_idx in linhas_a_usadas:
                record_candidate(
                    b_idx,
                    a_idx,
                    100.0,
                    "EXATO_BLOQUEADO",
                    eligible=True,
                    blocked=True,
                    reason="Referência já utilizada por outro nome.",
                )
                continue
            record_candidate(b_idx, a_idx, 100.0, "EXATO_NORMALIZADO", eligible=True)
            assign(b_idx, a_idx, 100.0, "EXATO_NORMALIZADO")
            break
    report(0.2)

    identity_lookup: dict[tuple[str, str], list[int]] = {}
    reference_identities = []
    for a_idx, nome_a_norm in enumerate(lista_a_normalizada):
        identity = _identidade_nome(nome_a_norm)
        reference_identities.append(identity)
        if identity:
            identity_lookup.setdefault(identity, []).append(a_idx)

    for b_idx, nome_b_norm in enumerate(nomes_b_normalizados):
        if b_idx % 128 == 0:
            check_cancel(should_cancel)
            report(0.2 + 0.2 * b_idx / total)
        if not nome_b_norm or b_idx in linhas_b_atribuidas:
            continue
        tq = set(tokens_significativos(nome_b_norm))
        identity_candidates = []
        for candidate_idx, a_idx in enumerate(identity_lookup.get(_identidade_nome(nome_b_norm), [])):
            if candidate_idx % 128 == 0:
                check_cancel(should_cancel)
            if a_idx in linhas_a_usadas:
                if match_details is not None:
                    score = score_par(nome_b_norm, lista_a_normalizada[a_idx], tq, lista_a_tokens[a_idx])
                    record_candidate(
                        b_idx,
                        a_idx,
                        score,
                        "IDENTIDADE_BLOQUEADA",
                        eligible=score >= threshold,
                        blocked=True,
                        reason="Referência já utilizada por outro nome.",
                    )
                continue
            score = score_par(nome_b_norm, lista_a_normalizada[a_idx], tq, lista_a_tokens[a_idx])
            if score < threshold:
                record_candidate(b_idx, a_idx, score, "MESMA_IDENTIDADE_ABAIXO_THRESHOLD")
                continue
            record_candidate(b_idx, a_idx, score, "MESMA_IDENTIDADE", eligible=True)
            identity_candidates.append((score, a_idx))
        if identity_candidates:
            score, a_idx = max(identity_candidates, key=lambda item: (item[0], -item[1]))
            assign(b_idx, a_idx, score, "MESMA_IDENTIDADE_MAIOR_SCORE")
    report(0.4)

    for b_idx, nome_b_norm in enumerate(nomes_b_normalizados):
        if b_idx % 64 == 0:
            check_cancel(should_cancel)
            report(0.4 + 0.6 * b_idx / total)
        if not nome_b_norm or b_idx in linhas_b_atribuidas:
            continue

        tq = set(tokens_significativos(nome_b_norm))
        query_identity = _identidade_nome(nome_b_norm)
        best_score = 0.0
        for a_idx, nome_a_norm in enumerate(lista_a_normalizada):
            if a_idx % 256 == 0:
                check_cancel(should_cancel)
            if a_idx in linhas_a_usadas and not track_best_scores and match_details is None:
                continue
            score = score_par(nome_b_norm, nome_a_norm, tq, lista_a_tokens[a_idx])
            if best_scores is not None:
                best_score = max(best_score, score)
            if a_idx in linhas_a_usadas:
                record_candidate(
                    b_idx,
                    a_idx,
                    score,
                    "BLOQUEADO_USADO",
                    eligible=score >= threshold,
                    blocked=True,
                    reason="Referência já utilizada por outro nome.",
                )
                continue
            if score < threshold:
                record_candidate(b_idx, a_idx, score, "ABAIXO_THRESHOLD")
                continue
            same_identity = (
                query_identity is not None
                and query_identity == reference_identities[a_idx]
            )
            if min_cross_identity_score is not None and not same_identity:
                if (
                    len(tq) < 2
                    or len(lista_a_tokens[a_idx]) < 2
                    or not _identidades_compativeis(
                        query_identity,
                        reference_identities[a_idx],
                        min_cross_identity_score,
                    )
                ):
                    record_candidate(
                        b_idx,
                        a_idx,
                        score,
                        "BLOQUEADO_REGRA",
                        blocked=True,
                        reason="Fuzzy entre identidades diferentes exige primeiro e último token compatíveis.",
                    )
                    continue
                if score < min_cross_identity_score:
                    record_candidate(
                        b_idx,
                        a_idx,
                        score,
                        "BLOQUEADO_REGRA",
                        blocked=True,
                        reason=f"Score abaixo do mínimo conservador de {min_cross_identity_score:.1f}.",
                    )
                    continue
            record_candidate(b_idx, a_idx, score, "FUZZY_GLOBAL", eligible=True)
            propostas.append((score, b_idx, a_idx))
        if best_scores is not None:
            best_scores[b_idx] = round(best_score, 1)

    propostas.sort(key=lambda item: (-item[0], item[1], item[2]))

    for score, b_idx, a_idx in propostas:
        if b_idx in linhas_b_atribuidas:
            if match_details is not None:
                conflicts[b_idx].append({
                    "a_idx": a_idx,
                    "score": round(score, 1),
                    "winner_b_idx": b_idx,
                    "winner_score": assigned_scores[b_idx],
                    "reason": "O nome já recebeu um candidato com score maior.",
                })
            continue
        if a_idx in linhas_a_usadas:
            if match_details is not None:
                winner_b_idx = assigned_by_a.get(a_idx)
                conflicts[b_idx].append({
                    "a_idx": a_idx,
                    "score": round(score, 1),
                    "winner_b_idx": winner_b_idx,
                    "winner_score": (
                        assigned_scores[winner_b_idx]
                        if winner_b_idx is not None else None
                    ),
                    "reason": "A referência foi atribuída a outro nome com score maior.",
                })
            continue
        assign(b_idx, a_idx, score, "FUZZY_GLOBAL")

    if match_details is not None:
        for b_idx, detail in enumerate(match_details):
            detail.update({
                "assignment_index": assignment_indices[b_idx],
                "stage": (
                    assigned_stage[b_idx]
                    if assignment_indices[b_idx] >= 0
                    else ("SEM_MATCH_CONFLITO" if conflicts[b_idx] else "SEM_MATCH")
                ),
                "score": assigned_scores[b_idx],
                "candidates_considered": candidate_considered[b_idx],
                "candidates_eligible": candidate_eligible[b_idx],
                "candidates_blocked": candidate_blocked[b_idx],
                "top_candidates": candidate_top[b_idx],
                "conflicts": conflicts[b_idx],
            })

    check_cancel(should_cancel)
    report(1.0)
    return assignment_indices, best_scores


def atribuir_indices_matches(
    nomes_b_normalizados: list[str],
    lista_a_normalizada: list[str],
    lista_a_tokens: list[set[str]],
    threshold: float,
    progress=None,
    should_cancel=None,
    min_cross_identity_score: float | None = None,
    match_details: list[dict] | None = None,
) -> list[int]:
    assignment_indices, _ = _calcular_atribuicoes(
        nomes_b_normalizados,
        lista_a_normalizada,
        lista_a_tokens,
        threshold,
        track_best_scores=False,
        progress=progress,
        should_cancel=should_cancel,
        min_cross_identity_score=min_cross_identity_score,
        match_details=match_details,
    )
    return assignment_indices


def atribuir_melhores_matches(
    nomes_b_originais: list[str],
    nomes_b_normalizados: list[str],
    lista_a_original: list[str],
    lista_a_normalizada: list[str],
    lista_a_tokens: list[set[str]],
    threshold: float,
    progress=None,
    should_cancel=None,
):
    nomes_finais = list(nomes_b_originais)
    statuses = [EMPTY_NAME if not nome_norm else NOT_FOUND for nome_norm in nomes_b_normalizados]
    assignment_indices, scores = _calcular_atribuicoes(
        nomes_b_normalizados,
        lista_a_normalizada,
        lista_a_tokens,
        threshold,
        progress=progress,
        should_cancel=should_cancel,
    )
    assert scores is not None
    for b_idx, a_idx in enumerate(assignment_indices):
        if a_idx < 0:
            continue
        nomes_finais[b_idx] = lista_a_original[a_idx]
        statuses[b_idx] = MATCH_FOUND

    return nomes_finais, statuses, scores
