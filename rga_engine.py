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

# v3 (template atualizado, set/2026): a planilha-base trocou "PERCENTUAL DE
# EXECUÇÃO FINANCEIRA" e "META [GERAL|ESPECÍFICA/AÇÃO] PACTUADA" por novos
# campos — "INDICADOR DA META GERAL"/"FONTE DOS DADOS" na Meta Geral, e
# "A AQUISIÇÃO/CONTRATAÇÃO FOI PREVISTA EM PLANO DE APLICAÇÃO?"/"HÁ VEDAÇÃO
# EXPRESSA DO ITEM ADQUIRIDO?" na Meta Específica. Os dois campos novos de
# cada bloco ainda não têm extração implementada (ver `indicador`/`fonte`
# em MG_COLS e `aquisicao_prevista`/`vedacao_expressa` em ME_COLS) — o RGA
# de referência usado até agora é de um formulário mais antigo que não os
# traz, então ficam "" até termos um PDF de exemplo com esses campos.
MG_COLS = {
    "descricao":         1,
    "polaridade":        2,
    "indicador":         3,
    "fonte":             4,
    "sinesp":            5,
    "val_ref_plano":     6,
    "val_ref_monitorado":7,
    "val_alcance":       8,
    "resultado":         9,
}

ME_COLS = {
    "descricao":          1,
    "bens":               2,
    "aquisicao_prevista": 3,
    "vedacao_expressa":   4,
    "polaridade":         5,
    "indicador":          6,
    "fonte":              7,
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
    """Lê os valores 'TOTAL DISPONIBILIZADO' e 'EXEC. FINANCEIRO TOTAL' do
    topo do relatório (página 1), localizando o rótulo pela sequência exata
    de palavras e pegando o valor 'R$...' logo abaixo dele — evita depender
    de coordenadas fixas ou da ordem em que o texto corrido aparece."""
    total_disp = _find_value_below_label(page0, ["TOTAL", "DISPONIBILIZADO"]) or ""
    exec_val = _find_value_below_label(page0, ["EXEC.", "FINANCEIRO"]) or ""

    # Recurso final (bem menos confiável): texto corrido com layout.
    if not total_disp or not exec_val:
        t = page0.extract_text(layout=True) or ""
        if not total_disp:
            m = re.search(r"TOTAL DISPONIBILIZADO.*?(R\$[\d\.,]+)", t, re.DOTALL)
            total_disp = m.group(1) if m else ""
        if not exec_val:
            m = re.search(r"EXEC\.?\s*FINANCEIRO TOTA\w*\s*\n?\s*(R\$[\d\.,]+)", t)
            exec_val = m.group(1) if m else ""

    return total_disp, exec_val


def _find_page_with_text(pdf, needle, start=0):
    """Procura a primeira página (a partir de `start`) cujo texto contém
    `needle`. Páginas intermediárias variam de quantidade dependendo do
    conteúdo condicional do RGA (ex.: caixa de alerta de divergência
    financeira), então o número de página nunca deve ser fixo — apenas a
    presença de um texto-âncora estável é confiável."""
    for pi in range(start, len(pdf.pages)):
        text = pdf.pages[pi].extract_text() or ""
        if needle in text:
            return pi
    return None


def _extract_meta_geral(pdf):
    p1 = pdf.pages[0]
    t1 = p1.extract_text(layout=True) or ""

    # A descrição/polaridade da Meta Geral normalmente está na página 0, mas
    # quando o RGA traz a caixa de alerta "Executado da Gestão MENOR/MAIOR
    # que o Executado Financeiro", esse bloco inteiro (descrição+polaridade)
    # é empurrado para a página seguinte — por isso a busca é dinâmica.
    desc_page_idx = _find_page_with_text(pdf, "META GERAL DO PLANO DE APLICAÇÃO")
    desc_page = pdf.pages[desc_page_idx] if desc_page_idx is not None else p1
    desc_text = desc_page.extract_text(layout=True) or t1

    # Os campos numéricos (1.1 a 1.5) ficam em uma página própria, cujo
    # índice também varia pelo mesmo motivo — nunca assumir pdf.pages[1].
    meta_page_idx = _find_page_with_text(pdf, "Avaliação da Meta Geral")
    p2 = pdf.pages[meta_page_idx] if meta_page_idx is not None else None
    words2 = p2.extract_words() if p2 else []
    t2 = p2.extract_text(layout=True) if p2 else ""

    # Descrição da Meta Geral
    desc_m = re.search(
        r"META GERAL DO PLANO DE APLICAÇÃO\s*\n(.+?)(?:\n\s*Polaridade)", desc_text, re.DOTALL
    )
    descricao = _clean(desc_m.group(1)) if desc_m else ""

    total_disp, exec_val = _extract_visao_geral_financeira(p1)

    pol_m = re.search(r"Polaridade do Indicador:.*?(Quanto \w+, \w+)", desc_text)
    polaridade = pol_m.group(1) if pol_m else ""

    sinesp = ""
    sin_label = _seq_pos(words2, ["Sinesp?"])
    if sin_label is not None:
        top_bound = words2[sin_label]["top"]
        bottom_bound = top_bound + 20
        sim_pos = next((w for w in words2 if w["text"] == "Sim"
                         and top_bound <= w["top"] <= bottom_bound), None)
        nao_pos = next((w for w in words2 if w["text"] == "Não"
                         and top_bound <= w["top"] <= bottom_bound), None)
        rows = []
        if sim_pos:
            rows.append(("SIM", sim_pos["x0"]))
        if nao_pos:
            rows.append(("NÃO", nao_pos["x0"]))
        marker_x0 = None
        for c in p2.curves:
            if not c.get("fill"):
                continue
            color = c.get("non_stroking_color") or (1, 1, 1)
            if all(v > 0.95 for v in color):
                continue
            w = c["x1"] - c["x0"]
            h = c["bottom"] - c["top"]
            if w < 10 and h < 10 and top_bound - 5 <= c["top"] <= bottom_bound:
                marker_x0 = c["x0"]
                break
        if marker_x0 is not None and rows:
            sinesp = min(rows, key=lambda r: abs(r[1] - marker_x0))[0]

    # 1.3. Valores da Meta Geral — checkbox "Plano não possui Meta Geral
    # mensurável". Quando marcado, os 4 campos numéricos abaixo (Meta
    # Pactuada / Valores de Referência / Valor-Alcance) e o alcance
    # automático (1.4) não se aplicam; em seu lugar, o campo de texto livre
    # "Demonstre se o objetivo está sendo alcançado (obrigatório)" é
    # preenchido — ver _write_meta_geral, que mescla essas colunas.
    sem_meta_mensuravel = _checkbox_marked(
        words2, ["Plano", "não", "possui", "Meta", "Geral", "mensurável"]
    )

    demonstre = _extract_paragraph_after(
        t2, r"Demonstre se o objetivo está sendo alcançado \(obrigatório\)",
        [r"\n\s*1\.4\."],
    )
    observacoes = _extract_paragraph_after(
        t2, r"1\.5\. Observações complementares \(opcional\)",
        [r"\n\s*2\.\s"],
    )

    idx = _seq_pos(words2, ["Valor", "de", "Referência", "(Apresentado"])
    val_ref_p = _value_below(words2, idx)

    idx = _seq_pos(words2, ["Valor", "de", "Referência", "(Monitorado"])
    val_ref_mon = _value_below(words2, idx)

    idx = _seq_pos(words2, ["Valor/Alcance", "da", "Meta", "Geral"])
    val_alc = _value_below(words2, idx)

    resultado = ""
    marco_14 = _word_top(p2, "1.4.", contains=False) if p2 else None
    marco_15 = _word_top(p2, "1.5.", contains=False) if p2 else None
    if marco_14:
        top_bound = marco_14[1]
        bottom_bound = marco_15[1] if marco_15 else top_bound + 110
        rows = _alcance_option_rows(p2, top_bound, bottom_bound)
        if rows:
            sel = _detect_selected_option(p2, top_bound, bottom_bound, rows)
            if sel:
                resultado = sel

    # TODO: "Indicador da Meta Geral" e "Fonte dos Dados" são campos novos
    # do template (set/2026) que ainda não têm extração implementada — o
    # RGA de referência usado até agora não os traz no formulário. Assim
    # que houver um PDF de exemplo com esses campos, localizar o rótulo
    # (provavelmente logo após a descrição/polaridade, no mesmo padrão já
    # usado em _extract_meta_especifica_avaliacao para Indicador/Fonte da
    # Meta Específica) e preencher aqui.
    indicador = ""
    fonte = ""

    return {
        "descricao": (
            f"META GERAL:  {descricao}"
            f"    VALOR TOTAL DISPONIBILIZADO  {total_disp}"
            f"    VALOR EXECUTADO FINANCEIRO TOTAL  {exec_val}"
            f"    STATUS  EM EXECUÇÃO"
        ),
        "polaridade":         polaridade,
        "indicador":          indicador,
        "fonte":              fonte,
        "sinesp":             sinesp,
        "val_ref_plano":      val_ref_p,
        "val_ref_monitorado": val_ref_mon,
        "val_alcance":        val_alc,
        "resultado":          resultado,
        "sem_meta_mensuravel": sem_meta_mensuravel,
        "demonstre":          demonstre,
        "observacoes":        observacoes,
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
    pequeno círculo preenchido com cor sólida (não branca); as demais têm
    só o contorno (ou um preenchimento branco de fundo, que é ignorado).
    `rows` é uma lista [(rótulo, top), ...] com a posição de cada opção."""
    marker_top = None
    for c in page.curves:
        if not c.get("fill"):
            continue
        color = c.get("non_stroking_color") or (1, 1, 1)
        if all(v > 0.95 for v in color):
            continue  # preenchimento branco de fundo, não é a marcação
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


def _seq_pos(words, seq, min_top=0):
    """Retorna o índice em `words` onde começa a sequência exata de tokens
    `seq` (ex.: ["Meta","Geral","Pactuada"]), a partir de min_top. Usado
    para localizar rótulos sem depender de regex sobre texto corrido."""
    n = len(seq)
    for i in range(len(words) - n + 1):
        if words[i]["top"] < min_top:
            continue
        if all(words[i + j]["text"] == seq[j] for j in range(n)):
            return i
    return None


def _checkbox_marked(words, seq, min_top=0):
    """Detecta se o checkbox (☑/☐) imediatamente antes da sequência de
    tokens `seq` está marcado. Usado tanto para 'Plano não possui Meta
    Geral mensurável' quanto para 'Meta Específica não possui indicador
    mensurável' — o símbolo do checkbox é sempre o token anterior ao
    rótulo, então basta localizar o rótulo e olhar um token pra trás."""
    idx = _seq_pos(words, seq, min_top=min_top)
    if idx is None or idx == 0:
        return False
    return words[idx - 1]["text"] == "☑"


def _extract_paragraph_after(text, start_pattern, end_patterns=()):
    """Captura o parágrafo de texto livre que aparece logo após um rótulo
    (`start_pattern`, regex), até o primeiro dos `end_patterns` (regex) que
    aparecer depois dele, ou até o fim do texto se nenhum aparecer. Usado
    para campos de texto livre (justificativas/observações) cujo conteúdo
    não tem tamanho fixo."""
    m = re.search(start_pattern + r"\s*\n", text)
    if not m:
        return ""
    rest = text[m.end():]
    end_idx = len(rest)
    for pat in end_patterns:
        em = re.search(pat, rest)
        if em and em.start() < end_idx:
            end_idx = em.start()
    return _clean(rest[:end_idx])


def _combine_justificativa(demonstre, observacoes):
    """Concatena o texto do campo obrigatório ('Demonstre se o objetivo
    está sendo alcançado') com o do campo opcional ('Observações
    complementares'), com um único espaço entre os dois quando ambos
    existem (e sem espaço sobrando quando só um deles existe)."""
    partes = [p.strip() for p in (demonstre, observacoes) if p and p.strip()]
    return " ".join(partes)


def _merge_row_range(ws, row, first_col, last_col, value):
    """Mescla as células de `first_col` a `last_col` na linha `row` e grava
    `value` na célula resultante. Remove antes qualquer mesclagem já
    existente que se sobreponha ao intervalo (ex.: as mesclagens H:I / J:K
    herdadas do template) — o openpyxl recusa criar uma mesclagem que
    sobreponha uma mesclagem já existente."""
    for rng in [r for r in list(ws.merged_cells.ranges)
                if r.min_row <= row <= r.max_row
                and not (r.max_col < first_col or r.min_col > last_col)]:
        ws.unmerge_cells(str(rng))
    ws.merge_cells(start_row=row, start_column=first_col,
                    end_row=row, end_column=last_col)
    ws.cell(row=row, column=first_col).value = value


def _value_below(words, label_idx, dx=(-15, 60), dy=(4, 25), span=1):
    """A partir da posição de um rótulo (índice em `words`), acha o token
    mais próximo posicionado logo abaixo dele (mesma coluna) — é assim que
    os campos de valor aparecem no PDF, um pouco abaixo do rótulo.

    `span` é quantos tokens (a partir de `label_idx`) fazem parte do
    próprio rótulo, usado como ponto de partida para achar a última linha
    do rótulo. Alguns rótulos são longos o bastante para quebrar em duas
    linhas (ex.: "...Exercício financeiro em" / "análise):"), e a
    continuação quebrada aparece alinhada à mesma margem esquerda do
    rótulo (mesmo x0), só que numa linha um pouco abaixo — por isso,
    depois do `span`, procuramos também por tokens alinhados a essa
    margem logo abaixo da última linha já conhecida e estendemos `lt`
    até eles, para não confundir a continuação do rótulo com o valor
    real (que vem depois, e não está alinhado a essa margem)."""
    if label_idx is None:
        return ""
    lx = words[label_idx]["x0"]
    lt = max(w["top"] for w in words[label_idx:label_idx + span])
    while True:
        wrapped = [
            w for w in words
            if lt < w["top"] <= lt + 15 and abs(w["x0"] - lx) <= 3
        ]
        if not wrapped:
            break
        new_lt = max(w["top"] for w in wrapped)
        if new_lt <= lt:
            break
        lt = new_lt
    best = None
    for w in words:
        if not (lt + dy[0] <= w["top"] <= lt + dy[1]):
            continue
        if not (lx + dx[0] <= w["x0"] <= lx + dx[1]):
            continue
        if best is None or w["top"] < best["top"]:
            best = w
    return best["text"] if best else ""


def _find_value_below_label(page, label_words, value_regex=r"^R\$[\d\.,]+$",
                             x_pad=(-10, 110), y_max=90, min_top=0):
    """Localiza a sequência de tokens `label_words` e devolve o primeiro
    token cujo texto bate com `value_regex`, posicionado abaixo do rótulo
    e na mesma coluna (x0 próximo) — usado para os cartões financeiros,
    que têm o valor exibido logo abaixo do título do cartão."""
    words = page.extract_words()
    idx = _seq_pos(words, label_words, min_top=min_top)
    if idx is None:
        return None
    label_x0, label_top = words[idx]["x0"], words[idx]["top"]
    best = None
    for w in words:
        if not re.match(value_regex, w["text"]):
            continue
        if w["top"] <= label_top or w["top"] - label_top > y_max:
            continue
        if not (label_x0 + x_pad[0] <= w["x0"] <= label_x0 + x_pad[1]):
            continue
        if best is None or w["top"] < best["top"]:
            best = w
    return best["text"] if best else None


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
    words = page.extract_words()

    # ─ Descrição (coluna esquerda, exclui a caixa de valores à direita) ─
    # Bordas dinâmicas: esquerda = posição real do título "META ESPECÍFICA N"
    # nesta página; direita = posição do rótulo "Total" (caixa de valores),
    # menos uma margem. Evita depender de coordenadas fixas de página.
    title_x0 = None
    for w in words:
        if w["text"] == "META" and abs(w["top"] - title_top) < 2:
            title_x0 = w["x0"]
            break
    left = (title_x0 - 10) if title_x0 is not None else 184

    total_word = _word_top(page, "Total", contains=False, min_top=title_top)
    right = (total_word[0] - 15) if total_word else 490

    plan_word = _word_top(page, "Planejado:", contains=False, min_top=title_top)
    desc_bottom = plan_word[1] if plan_word else title_top + 60
    desc_txt = _crop_text(page, (left, title_top + 6, right, desc_bottom - 1))
    desc = _clean(desc_txt)

    planejado_m = re.search(r"Planejado:\s*R?\$?\s*([\d\.,]+)", t)
    empenhado_m = re.search(r"Empenhado:\s*R?\$?\s*([\d\.,]+)", t)
    executado_m = re.search(r"Executado:\s*R?\$?\s*([\d\.,]+)", t)
    planejado = planejado_m.group(1) if planejado_m else ""
    empenhado = empenhado_m.group(1) if empenhado_m else ""
    executado = executado_m.group(1) if executado_m else ""

    # O template não tem mais uma coluna de "% de execução financeira", mas
    # o percentual continua sendo usado internamente como fallback para
    # decidir o "Resultado Alcançado" quando não dá pra ler a opção marcada
    # no rádio (ver _resultado_from_pct abaixo).
    pct_m = re.search(r"([\d,\.]+)%\s+([\d,\.]+)%\s*\n?\s*Saldo:", t)
    exec_pct_num = None
    if pct_m:
        try:
            exec_pct_num = float(pct_m.group(2).replace(",", "."))
        except ValueError:
            pass

    pol_m = re.search(r"Polaridade do Indicador:.*?(Quanto \w+, \w+)", t)
    polaridade = pol_m.group(1) if pol_m else ""

    sem_indicador = _checkbox_marked(
        words, ["Meta", "Específica", "não", "possui", "indicador", "mensurável"],
        min_top=title_top,
    )

    # ─ Indicador / Fonte (rótulos + crop de coluna) ─
    # A posição das colunas "Indicador da Meta Específica"/"Fonte dos
    # Dados" varia de relatório para relatório (não existe uma margem fixa
    # confiável — confirmado comparando o RGA AP|EVM|2023, onde essa
    # coluna começa perto de x=50, com o RGA AC|RMVI|2023, onde começa
    # perto de x=193). A âncora confiável é a posição real do PRÓPRIO
    # rótulo nesta página, então o crop usa ind_word[0]/fonte_word[0]
    # diretamente em vez de uma margem derivada do título ou fixa.
    ind_word = _word_top(page, "Indicador", contains=False, min_top=title_top)
    fonte_word = _word_top(page, "Fonte", contains=False, min_top=title_top)
    resultado_word = _word_top(page, "resultado", contains=False, min_top=title_top)

    # Borda inferior do crop de Indicador/Fonte: quando a meta TEM
    # indicador mensurável, logo abaixo vem a linha "Meta Específica/Ação
    # Pactuada... / Valor de Referência (Apresentado...)"; quando NÃO tem,
    # o PDF pula direto para "O resultado alcançado foi:". Sem esse limite
    # dinâmico, o crop antigo (que ia direto até "resultado") engolia as
    # linhas de valores junto com o Indicador/Fonte quando elas existiam —
    # esse era o bug visto no RGA AC|RMVI|2023.
    pactuada_idx = _seq_pos(words, ["Meta", "Específica/Ação", "Pactuada"], min_top=title_top)
    pactuada_top = words[pactuada_idx]["top"] if pactuada_idx is not None else None

    indicador = fonte = ""
    if ind_word and fonte_word:
        candidatos = [b for b in (pactuada_top, resultado_word[1] if resultado_word else None) if b]
        bottom0 = (min(candidatos) - 1) if candidatos else (ind_word[1] + 90)
        top0 = ind_word[1] + 9
        indicador = _clean(_crop_text(page, (ind_word[0] - 5, top0, fonte_word[0] - 15, bottom0)))
        fonte     = _clean(_crop_text(page, (fonte_word[0] - 5, top0, page.width - 20, bottom0)))

    # ─ Valores de referência (só existem no PDF quando a meta TEM
    #   indicador mensurável — quando o checkbox está marcado, o RGA
    #   substitui esse bloco pelo campo de texto livre de justificativa,
    #   ver `justificativa` mais abaixo) ─
    val_ref_plano = val_ref_monitorado = val_alcance = ""
    if pactuada_idx is not None:
        idx = _seq_pos(words, ["Valor", "de", "Referência", "(Apresentado"], min_top=title_top)
        val_ref_plano = _value_below(words, idx)

        idx = _seq_pos(words, ["Valor", "de", "Referência", "(Monitorado"], min_top=title_top)
        val_ref_monitorado = _value_below(words, idx, dy=(4, 30), span=8)

        idx = _seq_pos(words, ["Valor/Alcance", "da", "Meta", "Específica/Ação"], min_top=title_top)
        val_alcance = _value_below(words, idx)

    # ─ Resultado alcançado (opção assinalada) ─
    resultado = ""
    if sem_indicador:
        # Quando a meta não possui indicador mensurável, o resultado
        # correto é sempre "Não se aplica" — não depende de percentual de
        # execução nem da marcação visual (que pode falhar dependendo de
        # como o PDF foi gerado/renderizado).
        resultado = "Não se aplica"
    elif resultado_word:
        top_bound = resultado_word[1]
        bottom_bound = top_bound + 100
        rows = _alcance_option_rows(page, top_bound, bottom_bound)
        if rows:
            sel = _detect_selected_option(page, top_bound, bottom_bound, rows)
            if sel:
                resultado = sel
    if not resultado:
        resultado = _resultado_from_pct(exec_pct_num)

    # Justificativa (texto livre) — quando a meta não possui indicador
    # mensurável, o campo obrigatório "Demonstre se o objetivo está sendo
    # alcançado" é o parágrafo que aparece logo após a última opção do
    # rádio de resultado ("Não se aplica"), até o início da próxima Meta
    # Específica ou da seção seguinte do relatório.
    justificativa = ""
    if sem_indicador:
        justificativa = _extract_paragraph_after(
            t, r"Não se aplica",
            [r"\n\s*META\s+ESPECÍFICA\s+\d", r"\n\s*3\.\s"],
        )

    # TODO: "A aquisição/contratação foi prevista em Plano de Aplicação?" e
    # "Há vedação expressa do item adquirido?" não são campos que existem
    # no formulário do RGA — segundo a orientação da Mariana, devem ser
    # calculados comparando os itens de "6.1. Detalhamento dos Itens por
    # Meta Específica" com o que foi planejado. Ainda em aberto como
    # agregar esse resultado (por item) numa única célula por Meta
    # Específica — ver pergunta feita no chat antes de implementar.
    aquisicao_prevista = ""
    vedacao_expressa = ""

    return {
        "numero":             num,
        "polaridade":         polaridade,
        "indicador":          indicador,
        "fonte":              fonte,
        "aquisicao_prevista": aquisicao_prevista,
        "vedacao_expressa":   vedacao_expressa,
        "val_ref_plano":      val_ref_plano,
        "val_ref_monitorado": val_ref_monitorado,
        "val_alcance":        val_alcance,
        "resultado":          resultado,
        "sem_indicador":      sem_indicador,
        "justificativa":      justificativa,
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


def _detect_item_columns(pdf, first_pi, first_top):
    """Calibra os limites de coluna da tabela 'Item / Bem/Serviço' a partir
    da primeira linha de item real (linha começando com "1.") da primeira
    Meta Específica do relatório.

    A posição dessas colunas NÃO é estável entre RGAs diferentes — mesmo
    dois relatórios da mesma Secretaria/Fundo podem ter larguras de coluna
    bem diferentes (confirmado comparando duas versões reais do RGA
    AC|RMVI|2023: a tabela de itens de uma delas começa em x≈74, a da
    outra em x≈189). Por isso não há mais uma posição fixa de template —
    a calibração usa a própria primeira linha de dados como referência:

    1. Acha a palavra "ND" no cabeçalho da tabela (rótulo limpo, sem
       quebra) para saber onde termina a coluna de descrição do item.
    2. Acha a linha do item "1." e agrupa as palavras dessa linha (mais a
       linha logo abaixo, já que valores como "R$2.503.740,00" e o nome
       do órgão costumam quebrar em duas linhas) que ficam à direita da
       coluna ND — cada uma dessas colunas (ND, 3 valores monetários, %
       Exec, Status, Ano Exec., Itens Adquiridos) aparece como um único
       token curto por linha, então agrupar por proximidade horizontal
       (>15pt de distância = nova coluna) separa as 8 colunas corretamente
       sem ser afetado pela descrição do item (que fica à esquerda de ND
       e é ignorada aqui).

    Se não for possível calibrar (relatório sem essa seção, ou formato
    inesperado), cai de volta nos valores padrão do template antigo."""
    page = pdf.pages[first_pi]
    words = page.extract_words()

    nd_x0 = None
    for w in words:
        if w["text"] == "ND" and first_top - 5 <= w["top"] <= first_top + 60:
            nd_x0 = w["x0"]
            break

    item_top = None
    item_x0 = None
    limit_x0 = nd_x0 if nd_x0 is not None else 250
    for w in words:
        if w["top"] > first_top + 5 and re.fullmatch(r"1\.", w["text"]) and w["x0"] < limit_x0:
            item_top = w["top"]
            item_x0 = w["x0"]
            break

    if item_top is None or nd_x0 is None:
        return _ITEM_COL_DEFAULTS

    band = [
        w for w in words
        if item_top - 1 <= w["top"] <= item_top + 35 and w["x0"] >= nd_x0 - 20
    ]
    xs = sorted(set(round(w["x0"]) for w in band))
    clusters = []
    for x in xs:
        if not clusters or x - clusters[-1] > 15:
            clusters.append(x)

    # Precisamos de 8 colunas depois de "item" (nd, 3 valores, % exec,
    # status, ano exec, itens adquiridos); sem elas todas, não dá pra
    # confiar na calibração — melhor usar o padrão do template.
    if len(clusters) < 8:
        return _ITEM_COL_DEFAULTS

    margin = 4
    cols = [item_x0 - margin, nd_x0 - margin]
    cols += [c - margin for c in clusters[1:8]]
    cols.append(page.width)
    return cols


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


def _extract_items_for_range(pdf, start_page, start_top, end_page, end_top, cols=None):
    """Extrai todos os itens (linhas) da tabela 'Item / Bem/Serviço' entre
    (start_page, start_top) e (end_page, end_top), usando crops de coluna
    por posição para não misturar o conteúdo de colunas vizinhas.

    `cols` são os limites de coluna já calibrados para este relatório (ver
    `_detect_item_columns`); se omitido, usa os valores padrão do
    template."""
    cols = cols if cols is not None else _ITEM_COL_DEFAULTS
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
                # -0.5pt de folga: `top` é arredondado (round(w["top"], 1))
                # pra agrupar palavras na mesma linha, mas o crop de coluna
                # (passo 2, mais abaixo) usa esse valor como topo exato da
                # banda — se o arredondamento subir o valor um tico acima
                # do top real da palavra (ex.: 199.598 → 199.6), o
                # within_bbox exclui a própria primeira linha do item
                # (número + primeira palavra da descrição) da banda. A
                # folga evita esse corte.
                item_starts.append((pi, top - 0.5, m.group(1)))

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

        # Coluna "Itens Adquiridos": texto livre (pode ter várias linhas —
        # ano(s) em que o item foi de fato adquirido, seguido de detalhes
        # da aquisição, ou apenas "Nenhum" quando nada foi vinculado a este
        # item planejado). Mantém como texto corrido (não "_joined_clean",
        # que colaria as palavras sem espaço).
        itens_adquiridos = _clean(" ".join(col_texts["itens_adquiridos"]))

        items.append({
            "numero":            num,
            "descricao":         item_desc,
            "orgao":             orgao,
            "nd":                nd,
            "val_planejado":     val_planejado,
            "vl_empenhado":      vl_empenhado,
            "vol_executado":     vol_executado,
            "pct_exec":          pct_val,
            "status":            status,
            "ano_exec":          ano,
            "itens_adquiridos":  itens_adquiridos,
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


def _find_report_year(pdf):
    """Acha o ano do exercício financeiro do relatório, no cabeçalho da
    primeira página ('RELATÓRIO DE GESTÃO — EXERCÍCIO FINANCEIRO 2025').
    É esse ano — o da prestação de contas — que decide quais itens da
    coluna "Itens Adquiridos" (em 6.1) contam como efetivamente
    adquiridos nesse RGA."""
    if not pdf.pages:
        return ""
    text = pdf.pages[0].extract_text() or ""
    m = re.search(r"EXERC[ÍI]CIO\s+FINANCEIRO\s+(20\d{2})", text, re.IGNORECASE)
    return m.group(1) if m else ""


def _extract_bens_por_meta(pdf):
    """Localiza a seção '6.1. Detalhamento dos Itens por Meta Específica' e
    extrai, para cada Meta Específica, a lista de itens efetivamente
    adquiridos NO ANO DA PRESTAÇÃO DE CONTAS (o "Exercício Financeiro" do
    cabeçalho do RGA) — conforme a própria coluna "Itens Adquiridos" da
    tabela, que traz o ano da aquisição (ou "Nenhum" quando o item
    planejado não teve nenhum bem/serviço vinculado a ele). Antes disso,
    calibra as colunas da tabela a partir da primeira linha de item da
    primeira Meta Específica (ver `_detect_item_columns`), já que a
    posição das colunas varia entre relatórios."""
    idx_page = None
    for pi, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        if "Detalhamento dos Itens por Meta Específica" in text:
            idx_page = pi
            break
    if idx_page is None:
        return {}

    starts, end = _find_meta_especifica_table_starts(pdf, idx_page)
    if not starts:
        return {}

    ano_exercicio = _find_report_year(pdf)
    first_num, first_pi, first_top = starts[0]
    cols = _detect_item_columns(pdf, first_pi, first_top)

    bens_por_meta = {}
    for i, (num, pi, top) in enumerate(starts):
        if i + 1 < len(starts):
            npi, ntop = starts[i + 1][1], starts[i + 1][2]
        else:
            npi, ntop = end
        items = _extract_items_for_range(pdf, pi, top, npi, ntop, cols=cols)
        if ano_exercicio:
            # Fonte da verdade: a própria coluna "Itens Adquiridos" traz o
            # ano em que o item foi adquirido (ou "Nenhum"). Só entram os
            # itens cujo ano bate com o exercício financeiro deste RGA.
            # Um item começando com "Nenhum" nunca conta como adquirido —
            # mesmo que o crop tenha vazado texto do item seguinte pra
            # dentro da mesma célula (linhas próximas na tabela), o que
            # faria o ano aparecer como substring mesmo sem pertencer a
            # este item.
            def _foi_adquirido(it):
                txt = (it.get("itens_adquiridos") or "").strip()
                if txt.startswith("Nenhum"):
                    return False
                return ano_exercicio in txt

            adquiridos = [it for it in items if _foi_adquirido(it)]
        else:
            # Não foi possível achar o ano do exercício no cabeçalho —
            # volta pro critério antigo (valor executado/empenhado > 0)
            # como salvaguarda, em vez de não trazer nada.
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


# Colunas da Meta Geral que, quando "Plano não possui Meta Geral
# mensurável" está marcado, são mescladas numa única célula com o texto
# de justificativa (ver _write_meta_geral). Corresponde às colunas B a I
# do template atual (Polaridade, Indicador, Fonte, Sinesp, Valor Ref.
# Apresentado/Monitorado, Valor/Alcance e Resultado Alcançado) — conforme
# orientação da Mariana: "AS COLUNAS B3 ATÉ I3 DEVEM SER MESCLADAS".
_MG_MERGE_KEYS = ("polaridade", "indicador", "fonte", "sinesp",
                   "val_ref_plano", "val_ref_monitorado",
                   "val_alcance", "resultado")

# Colunas equivalentes na tabela de Meta Específica (C a K), conforme
# orientação da Mariana: "AS COLUNAS 7C ATÉ 7K DEVEM SER MESCLADAS".
_ME_MERGE_KEYS = ("aquisicao_prevista", "vedacao_expressa", "polaridade",
                   "indicador", "fonte", "val_ref_plano",
                   "val_ref_monitorado", "val_alcance", "resultado")


def _write_meta_geral(ws, mg, tmpl_ws):
    for r_src, r_dst in [(MG_TITLE_ROW, MG_TITLE_ROW),
                         (MG_HEADER_ROW, MG_HEADER_ROW),
                         (MG_DATA_ROW,   MG_DATA_ROW),
                         (MG_GAP_ROW,    MG_GAP_ROW)]:
        _copy_row(tmpl_ws, r_src, ws, r_dst)

    for r in [MG_TITLE_ROW, MG_HEADER_ROW, MG_DATA_ROW]:
        _copy_merged_ranges(tmpl_ws, r, ws, r)

    sem_mensuravel = mg.get("sem_meta_mensuravel")
    for key, col in MG_COLS.items():
        if sem_mensuravel and key in _MG_MERGE_KEYS:
            continue
        ws.cell(row=MG_DATA_ROW, column=col).value = mg.get(key, "")

    if sem_mensuravel:
        texto = _combine_justificativa(mg.get("demonstre"), mg.get("observacoes"))
        _merge_row_range(ws, MG_DATA_ROW, MG_COLS["polaridade"],
                          MG_COLS["resultado"], texto)


def _write_meta_especifica(ws, meta, block_idx, tmpl_ws):
    base = ME_FIRST_TITLE_ROW + (block_idx - 1) * ME_BLOCK_HEIGHT
    tmpl_base = ME_FIRST_TITLE_ROW

    for offset in range(ME_BLOCK_HEIGHT):
        r_src = tmpl_base + offset
        r_dst = base + offset
        _copy_row(tmpl_ws, r_src, ws, r_dst)
        _copy_merged_ranges(tmpl_ws, r_src, ws, r_dst)

    dr = base + 2
    sem_indicador = meta.get("sem_indicador")
    for key, col in ME_COLS.items():
        if sem_indicador and key in _ME_MERGE_KEYS:
            continue
        ws.cell(row=dr, column=col).value = meta.get(key, "")

    if sem_indicador:
        texto = _combine_justificativa(meta.get("justificativa"), meta.get("observacoes"))
        _merge_row_range(ws, dr, ME_COLS["aquisicao_prevista"],
                          ME_COLS["resultado"], texto)


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
    if meta_geral.get("sem_meta_mensuravel"):
        # As colunas B a I (Polaridade, Indicador, Fonte, Sinesp, valores
        # e Resultado) foram mescladas com o texto de justificativa
        # ("Demonstre se o objetivo está sendo alcançado" + "Observações
        # complementares") — checar esse texto em vez de cada campo
        # individual, que fica em branco propositalmente nesse caso.
        chk("Meta Geral › Demonstre objetivo alcançado",
            _combine_justificativa(meta_geral.get("demonstre"), meta_geral.get("observacoes")))
    else:
        chk("Meta Geral › Polaridade",         meta_geral.get("polaridade"))
        # Indicador/Fonte da Meta Geral: campos novos do template, extração
        # ainda não implementada (ver TODO em _extract_meta_geral) — ficam
        # sempre "" por ora, e por isso sempre aparecem aqui até serem
        # implementados. Isso é esperado, não um bug.
        chk("Meta Geral › Indicador",          meta_geral.get("indicador"))
        chk("Meta Geral › Fonte",              meta_geral.get("fonte"))
        chk("Meta Geral › Sinesp",             meta_geral.get("sinesp"))
        chk("Meta Geral › Valor Ref. Plano",   meta_geral.get("val_ref_plano"))
        chk("Meta Geral › Valor Ref. Mon.",    meta_geral.get("val_ref_monitorado"))
        chk("Meta Geral › Valor/Alcance",      meta_geral.get("val_alcance"))
        chk("Meta Geral › Resultado Alcançado", meta_geral.get("resultado"))

    for m in metas:
        n = m.get("numero", "?")
        if m.get("sem_indicador"):
            # As colunas C a K (Aquisição Prevista, Vedação Expressa,
            # Polaridade, Indicador, Fonte, valores e Resultado) foram
            # mescladas com o texto de justificativa — checar esse texto
            # em vez de cada campo individual.
            chk(f"Meta {n} › Demonstre objetivo alcançado", m.get("justificativa"))
        else:
            chk(f"Meta {n} › Indicador",       m.get("indicador"))
            chk(f"Meta {n} › Fonte",           m.get("fonte"))
            # Aquisição prevista/Vedação expressa: campos novos do
            # template, extração ainda não implementada (ver TODO em
            # _extract_meta_especifica_avaliacao).
            chk(f"Meta {n} › Aquisição Prevista", m.get("aquisicao_prevista"))
            chk(f"Meta {n} › Vedação Expressa",   m.get("vedacao_expressa"))
            chk(f"Meta {n} › Val. Ref. Mon.",  m.get("val_ref_monitorado"))
            chk(f"Meta {n} › Valor/Alcance",   m.get("val_alcance"))
            chk(f"Meta {n} › Resultado",       m.get("resultado"))

    return missing
