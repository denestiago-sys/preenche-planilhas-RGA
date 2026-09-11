"""
rga_engine.py
Motor de extração do PDF do Relatório de Gestão (RGA/FNSP)
e geração da planilha Excel no formato do template.

v2: extração baseada em posição (bounding boxes / crops de coluna) em vez de
regex sobre o texto "corrido" da página. O RGA é um relatório com layout de
cartões e tabelas multi-coluna; o texto corrido (extract_text padrão) intercala
o conteúdo de colunas vizinhas e quebra números/valores no meio, o que fazia
o parser anterior perder a maioria dos itens adquiridos e trocar valores.
"""
import copy
import re
from io import BytesIO
from pathlib import Path

import pdfplumber
import openpyxl
from openpyxl.cell.cell import MergedCell

REQUIRED_TEMPLATE_NAME = "Planilha RGA(base).xlsx"

# ─── Estrutura do template ──────────────────────────────────────────────────
MG_TITLE_ROW  = 1
MG_HEADER_ROW = 2
MG_DATA_ROW   = 3
MG_GAP_ROW    = 4
ME_FIRST_TITLE_ROW = 5
ME_BLOCK_HEIGHT    = 4

MG_COLS = {
    "descricao":         1,
    "polaridade":        2,
    "exec_pct":          3,
    "sinesp":            4,
    "meta_pactuada":     5,
    "val_ref_plano":     6,
    "val_ref_monitorado":7,
    "val_alcance":       8,
    "alcance":           10,
}

ME_COLS = {
    "descricao":          1,
    "bens":               2,
    "polaridade":         3,
    "exec_pct":           4,
    "indicador":          5,
    "fonte":              6,
    "meta_pactuada":      7,
    "val_ref_plano":      8,
    "val_ref_monitorado": 9,
    "val_alcance":        10,
    "resultado":          11,
}

_ALCANCE_OPCOES = [
    "100% (Integral)", "75% a 99% (Alto)", "50% a 74% (Médio)",
    "Abaixo de 50% (Baixo)", "Não se aplica",
]


# ══════════════════════════════════════════════════════════════════════════════
# UTILITÁRIOS DE LEITURA DE PDF
# ══════════════════════════════════════════════════════════════════════════════

def extract_lines_from_rga_file(file_obj):
    """Extrai texto 'corrido' de cada página (usado apenas para checar se há
    texto selecionável no PDF enviado)."""
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
        open_arg = file_obj
    else:
        open_arg = str(file_obj)

    pages = []
    with pdfplumber.open(open_arg) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return pages


def _clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def _safe_bbox(page, x0, top, x1, bottom):
    """Normaliza e valida uma bbox contra os limites reais da página.
    Retorna None se a área resultante for inválida (zero ou negativa)."""
    x0 = max(0, min(x0, page.width))
    x1 = max(0, min(x1, page.width))
    top = max(0, min(top, page.height))
    bottom = max(0, min(bottom, page.height))
    if x1 - x0 < 1 or bottom - top < 1:
        return None
    return (x0, top, x1, bottom)


def _crop_text(page, bbox):
    """Extrai texto de uma região retangular (x0, top, x1, bottom) da página,
    limitando aos limites reais da página para evitar erro do pdfplumber."""
    safe = _safe_bbox(page, *bbox)
    if safe is None:
        return ""
    try:
        return page.within_bbox(safe).extract_text() or ""
    except Exception:
        return ""


def _crop_words(page, bbox):
    """Como _crop_text, mas devolve a lista de palavras (com posição)."""
    safe = _safe_bbox(page, *bbox)
    if safe is None:
        return []
    try:
        return page.within_bbox(safe).extract_words()
    except Exception:
        return []


def _word_top(page, text, contains=False, min_top=0):
    """Retorna (x0, top) da primeira ocorrência de uma palavra na página."""
    for w in page.extract_words():
        if w["top"] < min_top:
            continue
        if (contains and text in w["text"]) or (not contains and w["text"] == text):
            return w["x0"], w["top"]
    return None


def extract_rga_signature(pages):
    """Extrai UF, sigla e ano do RGA (ex: 'AP | EVM | 2023')."""
    for page in pages[:2]:
        m = re.search(r"([A-Z]{2})\s*\|\s*([A-Z0-9]+)\s*\|\s*(20\d{2})", page)
        if m:
            return {"uf": m.group(1), "sigla": m.group(2), "ano": int(m.group(3))}
    return {"uf": None, "sigla": None, "ano": None}


# ══════════════════════════════════════════════════════════════════════════════
# META GERAL (página 1 = visão geral financeira / página 2 = avaliação)
# ══════════════════════════════════════════════════════════════════════════════

def _extract_visao_geral_financeira(page0):
    """Lê os cartões financeiros do topo do relatório (página 1) via crop de
    coluna, usando os retângulos dos próprios cartões como referência —
    evita o problema de valores intercalados no texto corrido."""
    rects = page0.rects
    # Agrupa retângulos por banda vertical (linha de cartões)
    bands = {}
    for r in rects:
        h = r["bottom"] - r["top"]
        if 30 < h < 90 and (r["x1"] - r["x0"]) > 60:
            key = (round(r["top"]), round(r["bottom"]))
            bands.setdefault(key, []).append(r)

    total_disp = exec_val = ""
    # Ordena bandas por posição vertical; a banda com 4 cartões contendo
    # "TOTAL DISPONIBILIZADO" e "EXEC. FINANCEIRO" é a que interessa.
    for (top, bottom), rs in sorted(bands.items()):
        rs_sorted = sorted(rs, key=lambda r: r["x0"])
        if len(rs_sorted) < 3:
            continue
        texts = [
            _crop_text(page0, (r["x0"], top, r["x1"], bottom)) for r in rs_sorted
        ]
        joined = " ".join(texts)
        if "TOTAL DISPONIBILIZADO" in joined and "FINANCEIRO TOTA" in joined:
            for t in texts:
                if "TOTAL DISPONIBILIZADO" in t:
                    m = re.search(r"R\$[\d\.,]+", t)
                    total_disp = m.group(0) if m else ""
                if "FINANCEIRO TOTA" in t:
                    m = re.search(r"R\$[\d\.,]+", t)
                    exec_val = m.group(0) if m else ""
            break

    # Fallback: se os retângulos não seguirem o padrão esperado, tenta por
    # rótulo em texto corrido com layout preservado.
    if not total_disp or not exec_val:
        t = page0.extract_text(layout=True) or ""
        if not total_disp:
            m = re.search(r"TOTAL DISPONIBILIZADO.*?(R\$[\d\.,]+)", t, re.DOTALL)
            total_disp = m.group(1) if m else ""
        if not exec_val:
            m = re.search(r"EXEC\.?\s*FINANCEIRO TOTA\w*\s*\n?\s*(R\$[\d\.,]+)", t)
            exec_val = m.group(1) if m else ""

    return total_disp, exec_val


def _extract_meta_geral(pdf):
    p1 = pdf.pages[0]
    p2 = pdf.pages[1] if len(pdf.pages) > 1 else None
    t1 = p1.extract_text(layout=True) or ""
    t2 = p2.extract_text(layout=True) if p2 else ""

    # Descrição da Meta Geral
    desc_m = re.search(
        r"META GERAL DO PLANO DE APLICAÇÃO\s*\n(.+?)(?:\n\s*Polaridade)", t1, re.DOTALL
    )
    descricao = _clean(desc_m.group(1)) if desc_m else ""

    total_disp, exec_val = _extract_visao_geral_financeira(p1)

    pol_m = re.search(r"Polaridade do Indicador:.*?(Quanto \w+, \w+)", t2)
    polaridade = pol_m.group(1) if pol_m else ""

    exec_pct_m = re.search(r"([\d,\.]+%)\s*executado sobre o total disponibilizado", t2)
    exec_pct = exec_pct_m.group(1) if exec_pct_m else ""

    sin_m = re.search(r"utilizando dados do Sinesp\?\s*\n?\s*(Sim|Não)", t2)
    sinesp = sin_m.group(1).upper() if sin_m else ""

    pac_m = re.search(
        r"Meta Geral Pactuada\s+Valor de Referência \(Apresentado.*?\):\s*\n\s*(\S+)\s+(\S+)",
        t2,
    )
    meta_pac  = pac_m.group(1) if pac_m else ""
    val_ref_p = pac_m.group(2) if pac_m else ""

    mon_m = re.search(
        r"análise\):\s+Valor/Alcance da Meta Geral no exercício em análise \(%\):\s*\n\s*(\S+)\s+(\S+)",
        t2,
    )
    val_ref_mon = mon_m.group(1) if mon_m else ""
    val_alc     = mon_m.group(2) if mon_m else ""

    alcance = ""
    marco_14 = _word_top(p2, "1.4.", contains=False)
    marco_15 = _word_top(p2, "1.5.", contains=False)
    if marco_14:
        top_bound = marco_14[1]
        bottom_bound = marco_15[1] if marco_15 else top_bound + 110
        rows = _alcance_option_rows(p2, top_bound, bottom_bound)
        if rows:
            sel = _detect_selected_option(p2, top_bound, bottom_bound, rows)
            if sel:
                alcance = sel
    if not alcance:
        # Recurso: texto corrido (menos confiável — sempre "acha" a
        # primeira opção listada, já que as 5 opções aparecem sempre).
        sec14 = t2[t2.find("1.4."):] if "1.4." in t2 else ""
        for opt in _ALCANCE_OPCOES:
            if opt in sec14:
                alcance = opt
                break

    return {
        "descricao": (
            f"META GERAL:  {descricao}"
            f"    VALOR TOTAL DISPONIBILIZADO  {total_disp}"
            f"    VALOR EXECUTADO FINANCEIRO TOTAL  {exec_val}"
            f"    STATUS  EM EXECUÇÃO"
        ),
        "polaridade":         polaridade,
        "exec_pct":           exec_pct,
        "sinesp":             sinesp,
        "meta_pactuada":      meta_pac,
        "val_ref_plano":      val_ref_p,
        "val_ref_monitorado": val_ref_mon,
        "val_alcance":        val_alc,
        "alcance":            alcance,
    }


# ══════════════════════════════════════════════════════════════════════════════
# METAS ESPECÍFICAS — bloco de avaliação (título + indicador + polaridade)
# ══════════════════════════════════════════════════════════════════════════════

def _resultado_from_pct(pct):
    if pct is None:
        return "100% (Integral)"
    if pct >= 100:  return "100% (Integral)"
    if pct >= 75:   return "75% a 99% (Alto)"
    if pct >= 50:   return "50% a 74% (Médio)"
    return "Abaixo de 50% (Baixo)"


def _detect_selected_option(page, top_bound, bottom_bound, rows):
    """Localiza qual opção de rádio está marcada dentro da faixa vertical
    [top_bound, bottom_bound]. No PDF, a opção marcada é desenhada com um
    pequeno círculo preenchido (fill=True); as demais têm só o contorno.
    `rows` é uma lista [(rótulo, top), ...] com a posição de cada opção."""
    marker_top = None
    for c in page.curves:
        if not c.get("fill"):
            continue
        w = c["x1"] - c["x0"]
        h = c["bottom"] - c["top"]
        if w < 10 and h < 10 and top_bound <= c["top"] <= bottom_bound:
            marker_top = c["top"]
            break
    if marker_top is None:
        return None
    label, top = min(rows, key=lambda r: abs(r[1] - marker_top))
    return label if abs(top - marker_top) < 8 else None


def _alcance_option_rows(page, top_bound, bottom_bound):
    """Deduz a posição (top) de cada uma das 5 opções padrão de alcance
    (100%/75-99%/50-74%/Abaixo de 50%/Não se aplica) dentro da faixa vertical
    indicada, usando a primeira ('100%') e a última ('Não') como âncoras —
    evita depender de coordenadas fixas, que variam de relatório pra relatório."""
    first = last = None
    for w in page.extract_words():
        if not (top_bound <= w["top"] <= bottom_bound):
            continue
        if w["text"] == "100%":
            first = w["top"] if first is None else min(first, w["top"])
        if w["text"] == "Não":
            last = w["top"] if last is None else max(last, w["top"])
    if first is None or last is None:
        return None
    step = (last - first) / 4.0
    tops = [first + i * step for i in range(5)]
    return list(zip(_ALCANCE_OPCOES, tops))


def _find_meta_titles(pdf, limit_page=None):
    """Localiza cada card 'META ESPECÍFICA N' (título em caixa alta) e
    devolve [(numero, page_idx, top), ...] em ordem."""
    titles = []
    pages = pdf.pages if limit_page is None else pdf.pages[:limit_page]
    for pi, page in enumerate(pages):
        # Busca a sequência "META" "ESPECÍFICA" "<num>" em palavras consecutivas
        words = page.extract_words()
        for i in range(len(words) - 2):
            if (words[i]["text"] == "META" and words[i + 1]["text"] == "ESPECÍFICA"
                    and re.match(r"^\d+$", words[i + 2]["text"])):
                titles.append((words[i + 2]["text"], pi, words[i]["top"]))
    return titles


def _extract_meta_especifica_avaliacao(pdf, num, page_idx, title_top, next_top_page, next_top):
    page = pdf.pages[page_idx]
    t = page.extract_text(layout=True) or ""

    # ─ Descrição (coluna esquerda, exclui a caixa de valores à direita) ─
    desc = ""
    plan_word = _word_top(page, "Planejado:", contains=False, min_top=title_top)
    desc_bottom = plan_word[1] if plan_word else title_top + 60
    desc_txt = _crop_text(page, (184, title_top + 6, 490, desc_bottom - 1))
    desc = _clean(desc_txt)

    planejado_m = re.search(r"Planejado:\s*R?\$?\s*([\d\.,]+)", t)
    empenhado_m = re.search(r"Empenhado:\s*R?\$?\s*([\d\.,]+)", t)
    executado_m = re.search(r"Executado:\s*R?\$?\s*([\d\.,]+)", t)
    planejado = planejado_m.group(1) if planejado_m else ""
    empenhado = empenhado_m.group(1) if empenhado_m else ""
    executado = executado_m.group(1) if executado_m else ""

    pct_m = re.search(r"([\d,\.]+)%\s+([\d,\.]+)%\s*\n?\s*Saldo:", t)
    exec_pct_num = None
    exec_pct_str = ""
    if pct_m:
        exec_pct_str = pct_m.group(2).replace(".", ",") + "%"
        try:
            exec_pct_num = float(pct_m.group(2).replace(",", "."))
        except ValueError:
            pass

    pol_m = re.search(r"Polaridade do Indicador:.*?(Quanto \w+, \w+)", t)
    polaridade = pol_m.group(1) if pol_m else ""

    sem_indicador = "não possui indicador mensurável" in t

    # ─ Indicador / Fonte (crop em duas colunas) ─
    ind_word = _word_top(page, "Indicador", contains=False, min_top=title_top)
    fonte_word = _word_top(page, "Fonte", contains=False, min_top=title_top)
    resultado_word = _word_top(page, "resultado", contains=False, min_top=title_top)
    indicador = fonte = ""
    if ind_word and resultado_word:
        top0 = ind_word[1] + 9
        bottom0 = resultado_word[1] - 1
        fonte_x0 = fonte_word[0] - 3 if fonte_word else 375
        indicador = _clean(_crop_text(page, (184, top0, fonte_x0, bottom0)))
        fonte     = _clean(_crop_text(page, (fonte_x0, top0, 571, bottom0)))

    # ─ Resultado alcançado (opção assinalada) ─
    resultado = ""
    if resultado_word:
        top_bound = resultado_word[1]
        bottom_bound = top_bound + 100
        rows = _alcance_option_rows(page, top_bound, bottom_bound)
        if rows:
            sel = _detect_selected_option(page, top_bound, bottom_bound, rows)
            if sel:
                resultado = sel
    if not resultado:
        resultado = _resultado_from_pct(exec_pct_num)

    return {
        "numero":             num,
        "polaridade":         polaridade,
        "exec_pct":           exec_pct_str,
        "indicador":          indicador if not sem_indicador else indicador,
        "fonte":              fonte,
        "meta_pactuada":      "",
        "val_ref_plano":      "",
        "val_ref_monitorado": "",
        "val_alcance":        "",
        "resultado":          resultado,
        "bens":               [],
        "_desc":              desc,
        "_plan":              planejado,
        "_exec":              executado,
    }


def _extract_metas_especificas(pdf, section2_end_page):
    titles = _find_meta_titles(pdf, limit_page=section2_end_page + 1)
    metas = []
    for idx, (num, pi, top) in enumerate(titles):
        if idx + 1 < len(titles):
            next_pi, next_top = titles[idx + 1][1], titles[idx + 1][2]
        else:
            next_pi, next_top = pi, None
        m = _extract_meta_especifica_avaliacao(pdf, num, pi, top, next_pi, next_top)
        metas.append(m)

    seen, uniq = set(), []
    for m in metas:
        if m["numero"] not in seen:
            seen.add(m["numero"])
            uniq.append(m)
    uniq.sort(key=lambda x: int(x["numero"]))

    for m in uniq:
        m["descricao"] = (
            f"META ESPECÍFICA {m['numero']}:  {m.pop('_desc')}"
            f"    VALOR PLANEJADO  R$ {m.pop('_plan')}"
            f"    VALOR EXECUTADO  R$ {m.pop('_exec')}"
            f"    STATUS  EM EXECUÇÃO"
        )
    return uniq


# ══════════════════════════════════════════════════════════════════════════════
# 6.1 DETALHAMENTO DOS ITENS POR META ESPECÍFICA
# ══════════════════════════════════════════════════════════════════════════════
# Colunas da tabela de itens (posições em pontos, calibradas no template do
# relatório; ver comentário em _detect_item_columns).
_ITEM_COL_NAMES = [
    "item", "nd", "val_planejado", "vl_empenhado", "vol_executado",
    "pct_exec", "status", "ano_exec", "itens_adquiridos",
]
_ITEM_COL_DEFAULTS = [184, 292, 327, 366, 405, 443, 467, 497, 525, 584]


def _detect_item_columns(page):
    """Tenta calibrar os limites de coluna da tabela de itens a partir do
    cabeçalho 'Item / Bem/Serviço ND ... Status ...' desta página; se não
    encontrar, usa os valores padrão do template."""
    words = {w["text"]: w["x0"] for w in page.extract_words()
             if w["top"] < 700}
    header_hits = [k for k in ("Item", "ND", "Status") if k in words]
    if len(header_hits) < 2:
        return _ITEM_COL_DEFAULTS
    # Sem uma calibração completa e confiável, mantém o padrão do template
    # (validado empiricamente) — a detecção acima serve só para confirmar
    # que estamos de fato numa página de tabela de itens.
    return _ITEM_COL_DEFAULTS


def _header_row_bottom(page):
    """Localiza a linha de cabeçalho repetido da tabela de itens
    ('Item / Bem/Serviço ND ...') numa página de continuação, para poder
    excluí-la das colunas de dados."""
    for w in page.extract_words():
        if w["text"] == "Item" and w["x0"] < 250 and w["top"] < 700:
            return w["top"] + 10
    return None


def _find_meta_especifica_table_starts(pdf, start_page):
    """A partir da página do índice '6.1. Detalhamento...', localiza o
    início (page_idx, top) de cada tabela 'Meta Específica N —' e o ponto
    de término da seção (início de '6.2.' ou fim do documento)."""
    starts = []
    end = None
    for pi in range(start_page, len(pdf.pages)):
        page = pdf.pages[pi]
        text = page.extract_text() or ""
        if pi > start_page and re.search(r"6\.2\.|PENDÊNCIAS DAS CONTAS", text):
            end = (pi - 1, pdf.pages[pi - 1].height)
            break
        m = re.search(r"Meta Específica\s+(\d+)\s*[—-]", text)
        if m:
            # localiza o "top" exato do início do título nesta página
            top = None
            words = page.extract_words()
            for i, w in enumerate(words):
                if (w["text"] == "Meta" and i + 1 < len(words)
                        and words[i + 1]["text"] == "Específica"):
                    top = w["top"]
                    break
            starts.append((m.group(1), pi, top if top is not None else 0))
    if end is None:
        end = (len(pdf.pages) - 1, pdf.pages[-1].height)
    return starts, end


def _normalize_status(s):
    s = s.strip()
    if s.startswith("Aprovad"):
        return "Aprovada"
    if s.startswith("Cancelad"):
        return "Cancelada"
    if s.startswith("Reprovad"):
        return "Reprovada"
    return s


def _extract_items_for_range(pdf, start_page, start_top, end_page, end_top):
    """Extrai todos os itens (linhas) da tabela 'Item / Bem/Serviço' entre
    (start_page, start_top) e (end_page, end_top), usando crops de coluna
    por posição para não misturar o conteúdo de colunas vizinhas."""
    cols = _ITEM_COL_DEFAULTS
    col_ranges = list(zip(cols[:-1], cols[1:]))
    names = _ITEM_COL_NAMES

    # 1) Localiza as linhas de início de item ("N. ") na coluna Item, em
    #    todas as páginas do intervalo.
    item_starts = []  # (page_idx, top, numero)
    for pi in range(start_page, end_page + 1):
        page = pdf.pages[pi]
        top0 = start_top if pi == start_page else 0
        top1 = end_top if pi == end_page else page.height
        words = _crop_words(page, (col_ranges[0][0], top0, col_ranges[0][1], top1))
        lines = {}
        for w in words:
            lines.setdefault(round(w["top"], 1), []).append((w["x0"], w["text"]))
        for top in sorted(lines):
            toks = [t for _, t in sorted(lines[top])]
            line_text = " ".join(toks)
            m = re.match(r"^(\d+)\.\s", line_text)
            if m:
                item_starts.append((pi, top, m.group(1)))

    if not item_starts:
        return []

    # 2) Para cada item, delimita a banda (página/top inicial até o próximo
    #    item, ou fim do intervalo) e faz o crop de cada coluna nessa banda.
    items = []
    for i, (pi, top, num) in enumerate(item_starts):
        if i + 1 < len(item_starts):
            npi, ntop, _ = item_starts[i + 1]
        else:
            npi, ntop = end_page, end_top

        col_texts = {name: [] for name in names}
        for p in range(pi, npi + 1):
            page = pdf.pages[p]
            band_top = top if p == pi else 0
            band_bottom = ntop if p == npi else page.height
            band_top = max(0, band_top)
            band_bottom = min(page.height, band_bottom)
            if p != pi:
                # Página de continuação: pula o cabeçalho repetido da tabela
                # ("Item / Bem/Serviço ND ...") para não vazar na descrição.
                hdr_bottom = _header_row_bottom(page)
                if hdr_bottom is not None:
                    band_top = max(band_top, hdr_bottom)
            if band_bottom <= band_top:
                continue
            for name, (cx0, cx1) in zip(names, col_ranges):
                txt = _crop_text(page, (cx0, band_top, cx1, band_bottom))
                if txt:
                    col_texts[name].append(txt)

        item_desc = _clean(" ".join(col_texts["item"]))
        item_desc = re.sub(r"^\d+\.\s*", "", item_desc)
        # remove o nome do órgão executor (linha em caixa alta ao final)
        # e o guarda separadamente quando identificável
        orgao = ""
        m_org = re.search(
            r"\b(POLÍCIA (MILITAR|CIVIL)|SECRETARIA[ A-ZÀ-Ú]*|CORPO DE BOMBEIROS[ A-ZÀ-Ú]*)\b\s*$",
            item_desc,
        )
        if m_org:
            orgao = m_org.group(0).strip()
            item_desc = item_desc[: m_org.start()].strip()

        def _num(name):
            raw = "".join(col_texts[name]).replace("\n", "")
            raw = re.sub(r"[^\d,\.]", "", raw)
            return raw

        def _joined_clean(name):
            # Colunas de valor/rótulo curto: o PDF quebra o próprio token no
            # meio (ex.: "100." / "0%", "Investime" / "nto"), então NÃO se
            # insere espaço entre os fragmentos — apenas se remove a quebra
            # de linha, preservando a ordem em que os fragmentos aparecem.
            return re.sub(r"\s+", "", "".join(col_texts[name])).strip()

        nd_raw = _joined_clean("nd")
        nd = "Investimento" if nd_raw.startswith("Investime") else (
            "Custeio" if nd_raw.startswith("Custeio") else nd_raw
        )

        val_planejado = _num("val_planejado")
        vl_empenhado  = _num("vl_empenhado")
        vol_executado = _num("vol_executado")

        pct_raw = _joined_clean("pct_exec")
        pct_m = re.search(r"^([\d\.,]+)%", pct_raw)
        if not pct_m:
            pct_m = re.search(r"([\d\.,]+)%", pct_raw)
        pct_val = None
        if pct_m:
            try:
                pct_val = float(pct_m.group(1).rstrip(".,").replace(",", "."))
            except ValueError:
                pct_val = None

        status = _normalize_status(_joined_clean("status"))
        ano_raw = _joined_clean("ano_exec")
        ano_m = re.search(r"20\d{2}", ano_raw)
        ano = ano_m.group(0) if ano_m else ""

        items.append({
            "numero":        num,
            "descricao":     item_desc,
            "orgao":         orgao,
            "nd":            nd,
            "val_planejado": val_planejado,
            "vl_empenhado":  vl_empenhado,
            "vol_executado": vol_executado,
            "pct_exec":      pct_val,
            "status":        status,
            "ano_exec":      ano,
        })
    return items


def _format_bem_line(item):
    partes = [f"- {item['descricao']}"]
    if item.get("orgao"):
        partes.append(f"({item['orgao']})")
    detalhe = []
    if item.get("nd"):
        detalhe.append(item["nd"])
    if item.get("vol_executado"):
        detalhe.append(f"Executado: R$ {item['vol_executado']}")
    elif item.get("vl_empenhado"):
        detalhe.append(f"Empenhado: R$ {item['vl_empenhado']}")
    if item.get("pct_exec") is not None:
        detalhe.append(f"{item['pct_exec']:.1f}%".replace(".", ",") + " Exec")
    if item.get("ano_exec"):
        detalhe.append(f"Ano {item['ano_exec']}")
    if detalhe:
        partes.append("— " + "; ".join(detalhe))
    return " ".join(partes)


def _extract_bens_por_meta(pdf):
    """Localiza a seção '6.1. Detalhamento dos Itens por Meta Específica' e
    extrai, para cada Meta Específica, a lista de itens efetivamente
    adquiridos (Vol. Executado > 0 ou Vl. Empenhado > 0)."""
    idx_page = None
    for pi, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        if "Detalhamento dos Itens por Meta Específica" in text:
            idx_page = pi
            break
    if idx_page is None:
        return {}

    starts, end = _find_meta_especifica_table_starts(pdf, idx_page)
    bens_por_meta = {}
    for i, (num, pi, top) in enumerate(starts):
        if i + 1 < len(starts):
            npi, ntop = starts[i + 1][1], starts[i + 1][2]
        else:
            npi, ntop = end
        items = _extract_items_for_range(pdf, pi, top, npi, ntop)
        adquiridos = [
            it for it in items
            if (it["pct_exec"] or 0) > 0
            or _to_float(it["vol_executado"]) > 0
            or _to_float(it["vl_empenhado"]) > 0
        ]
        bens_por_meta[num] = [_format_bem_line(it) for it in adquiridos]
    return bens_por_meta


def _to_float(raw):
    if not raw:
        return 0.0
    try:
        return float(raw.replace(".", "").replace(",", "."))
    except ValueError:
        return 0.0


# ─── Entry point de extração ───────────────────────────────────────────────

def extract_rga_data(file_obj):
    """
    Retorna (meta_geral_dict, [meta_especifica_dict, ...]).
    Cada meta_especifica_dict tem chaves alinhadas a ME_COLS.
    """
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
        open_arg = file_obj
    else:
        open_arg = str(file_obj)

    with pdfplumber.open(open_arg) as pdf:
        if not pdf.pages:
            return None, []

        meta_geral = _extract_meta_geral(pdf)

        section2_end_page = len(pdf.pages) - 1
        for pi, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if "Resultados dos Indicadores por Meta Específica" in text:
                section2_end_page = pi
                break

        metas = _extract_metas_especificas(pdf, section2_end_page)
        bens  = _extract_bens_por_meta(pdf)
        for m in metas:
            m["bens"] = "\n\n".join(bens.get(m["numero"], []))

    return meta_geral, metas


# ══════════════════════════════════════════════════════════════════════════════
# GERAÇÃO DO EXCEL  (inalterado)
# ══════════════════════════════════════════════════════════════════════════════

def _copy_cell(src, dst):
    dst.value = src.value
    if src.has_style:
        dst.font          = copy.copy(src.font)
        dst.fill          = copy.copy(src.fill)
        dst.border        = copy.copy(src.border)
        dst.alignment     = copy.copy(src.alignment)
        dst.number_format = src.number_format


def _copy_row(ws_src, r_src, ws_dst, r_dst, max_col=11):
    for col in range(1, max_col + 1):
        src = ws_src.cell(row=r_src, column=col)
        dst = ws_dst.cell(row=r_dst, column=col)
        if isinstance(src, MergedCell):
            continue
        _copy_cell(src, dst)
    ws_dst.row_dimensions[r_dst].height = ws_src.row_dimensions[r_src].height


def _copy_merged_ranges(ws_src, r_src, ws_dst, r_dst):
    shift = r_dst - r_src
    for rng in ws_src.merged_cells.ranges:
        if rng.min_row == r_src and rng.max_row == r_src:
            try:
                ws_dst.merge_cells(
                    start_row=r_dst,
                    start_column=rng.min_col,
                    end_row=r_dst,
                    end_column=rng.max_col,
                )
            except Exception:
                pass


def _write_meta_geral(ws, mg, tmpl_ws):
    for r_src, r_dst in [(MG_TITLE_ROW, MG_TITLE_ROW),
                         (MG_HEADER_ROW, MG_HEADER_ROW),
                         (MG_DATA_ROW,   MG_DATA_ROW),
                         (MG_GAP_ROW,    MG_GAP_ROW)]:
        _copy_row(tmpl_ws, r_src, ws, r_dst)

    for r in [MG_TITLE_ROW, MG_HEADER_ROW, MG_DATA_ROW]:
        _copy_merged_ranges(tmpl_ws, r, ws, r)

    for key, col in MG_COLS.items():
        ws.cell(row=MG_DATA_ROW, column=col).value = mg.get(key, "")


def _write_meta_especifica(ws, meta, block_idx, tmpl_ws):
    base = ME_FIRST_TITLE_ROW + (block_idx - 1) * ME_BLOCK_HEIGHT
    tmpl_base = ME_FIRST_TITLE_ROW

    for offset in range(ME_BLOCK_HEIGHT):
        r_src = tmpl_base + offset
        r_dst = base + offset
        _copy_row(tmpl_ws, r_src, ws, r_dst)
        _copy_merged_ranges(tmpl_ws, r_src, ws, r_dst)

    dr = base + 2
    for key, col in ME_COLS.items():
        val = meta.get(key, "")
        ws.cell(row=dr, column=col).value = val


def generate_rga_excel_bytes(template_path: Path, meta_geral: dict, metas: list) -> bytes:
    tmpl_wb = openpyxl.load_workbook(template_path)
    tmpl_ws = tmpl_wb.active

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Planilha1"

    for col_letter, dim in tmpl_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = dim.width

    _write_meta_geral(ws, meta_geral, tmpl_ws)

    for idx, meta in enumerate(metas, start=1):
        _write_meta_especifica(ws, meta, idx, tmpl_ws)

    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"
    ws.sheet_view.zoomScale = 100

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def get_missing_cells(meta_geral: dict, metas: list) -> list:
    missing = []

    def chk(label, value):
        if not str(value or "").strip():
            missing.append(label)

    chk("Meta Geral › Descrição",          meta_geral.get("descricao"))
    chk("Meta Geral › Polaridade",         meta_geral.get("polaridade"))
    chk("Meta Geral › % Execução",         meta_geral.get("exec_pct"))
    chk("Meta Geral › Sinesp",             meta_geral.get("sinesp"))
    chk("Meta Geral › Meta Pactuada",      meta_geral.get("meta_pactuada"))
    chk("Meta Geral › Valor Ref. Plano",   meta_geral.get("val_ref_plano"))
    chk("Meta Geral › Valor Ref. Mon.",    meta_geral.get("val_ref_monitorado"))
    chk("Meta Geral › Valor/Alcance",      meta_geral.get("val_alcance"))
    chk("Meta Geral › Alcance Pactuado",   meta_geral.get("alcance"))

    for m in metas:
        n = m.get("numero", "?")
        chk(f"Meta {n} › Indicador",       m.get("indicador"))
        chk(f"Meta {n} › Fonte",           m.get("fonte"))
        chk(f"Meta {n} › Val. Ref. Mon.",  m.get("val_ref_monitorado"))
        chk(f"Meta {n} › Valor/Alcance",   m.get("val_alcance"))
        chk(f"Meta {n} › Resultado",       m.get("resultado"))

    return missing
