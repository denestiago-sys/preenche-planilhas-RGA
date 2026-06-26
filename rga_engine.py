"""
rga_engine.py
Motor de extração do PDF do Relatório de Gestão (RGA/FNSP)
e geração da planilha Excel no formato do template.
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
# Bloco Meta Geral  : linha 1 (título), linha 2 (cabeçalho), linha 3 (dados), linha 4 (espaço)
# Bloco Meta Esp.   : linha N+0 (título), N+1 (cabeçalho), N+2 (dados), N+3 (espaço)
MG_TITLE_ROW  = 1
MG_HEADER_ROW = 2
MG_DATA_ROW   = 3
MG_GAP_ROW    = 4
ME_FIRST_TITLE_ROW = 5   # primeira Meta Específica começa na linha 5
ME_BLOCK_HEIGHT    = 4   # linhas por bloco de Meta Específica

# Colunas Meta Geral
MG_COLS = {
    "descricao":         1,   # A
    "polaridade":        2,   # B
    "exec_pct":          3,   # C
    "sinesp":            4,   # D
    "meta_pactuada":     5,   # E
    "val_ref_plano":     6,   # F
    "val_ref_monitorado":7,   # G
    "val_alcance":       8,   # H
    # col I = merged com H no template
    "alcance":           10,  # J (merged J:K)
}

# Colunas Meta Específica
ME_COLS = {
    "descricao":          1,   # A
    "bens":               2,   # B
    "polaridade":         3,   # C
    "exec_pct":           4,   # D
    "indicador":          5,   # E
    "fonte":              6,   # F
    "meta_pactuada":      7,   # G
    "val_ref_plano":      8,   # H
    "val_ref_monitorado": 9,   # I
    "val_alcance":        10,  # J
    "resultado":          11,  # K
}


# ══════════════════════════════════════════════════════════════════════════════
# EXTRAÇÃO DO PDF
# ══════════════════════════════════════════════════════════════════════════════

def extract_lines_from_rga_file(file_obj):
    """Extrai texto do PDF do RGA (aceita file-like object ou Path)."""
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


def extract_rga_signature(pages):
    """Extrai UF, sigla e ano do RGA (ex: 'RJ | RMVI | 2023')."""
    for page in pages[:2]:
        m = re.search(r"([A-Z]{2})\s*\|\s*([A-Z0-9]+)\s*\|\s*(20\d{2})", page)
        if m:
            return {"uf": m.group(1), "sigla": m.group(2), "ano": int(m.group(3))}
    return {"uf": None, "sigla": None, "ano": None}


# ─── Meta Geral ────────────────────────────────────────────────────────────

def _extract_meta_geral(pages):
    p1 = pages[0] if pages else ""
    p2 = pages[1] if len(pages) > 1 else ""

    # Descrição
    desc_m = re.search(
        r"META GERAL DO PLANO DE APLICAÇÃO\s*\n(.+?)(?:\nPolaridade)", p1, re.DOTALL
    )
    descricao = _clean(desc_m.group(1)) if desc_m else ""

    # Valores financeiros (sequência no PDF):
    # repasse_orig, suplementar, total_repassado, rendimentos, total_disp, exec_gg, ...
    all_vals = re.findall(r"\$([\d,\.]+)", p1)
    total_disp = all_vals[4] if len(all_vals) > 4 else ""
    exec_val   = all_vals[5] if len(all_vals) > 5 else ""

    pol_m = re.search(r"Polaridade do Indicador:[^\n]*(Quanto \w+, \w+)", p1)
    polaridade = pol_m.group(1) if pol_m else ""

    exec_pct_m = re.search(r"([\d,]+%)\s*executado sobre o total disponibilizado", p2)
    exec_pct = exec_pct_m.group(1) if exec_pct_m else ""

    sin_m = re.search(r"1\.2\..*?\n\s*(Sim|Não)", p2, re.DOTALL)
    sinesp = sin_m.group(1).upper() if sin_m else ""

    pac_m = re.search(r"Meta Geral Pactuada\b.*?\n\s*(\S+)\s+(\S+)", p2, re.DOTALL)
    meta_pac  = pac_m.group(1) if pac_m else ""
    val_ref_p = pac_m.group(2) if pac_m else ""

    mon_m = re.search(r"análise\):\s+Valor/Alcance[^\n]*\n\s*(\S+)\s+(\S+)", p2)
    if not mon_m:
        mon_m = re.search(r"análise\):\s*\n\s*(\S+)\s*\n\s*(\S+)", p2)
    val_ref_mon = mon_m.group(1) if mon_m else ""
    val_alc     = mon_m.group(2) if mon_m else ""

    alcance = ""
    sec14 = p2[p2.find("1.4."):]
    for opt in [
        "100% (Integral)", "75% a 99% (Alto)", "50% a 74% (Médio)",
        "Abaixo de 50% (Baixo)", "Não se aplica"
    ]:
        if opt in sec14:
            alcance = opt
            break

    return {
        "descricao": (
            f"META GERAL:  {descricao}"
            f"    VALOR TOTAL  R$ {total_disp}"
            f"    VALOR EXECUTADO  R$ {exec_val}"
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


# ─── Metas Específicas ─────────────────────────────────────────────────────

def _resultado_from_block(text):
    """Determina resultado. Justificativa presente = Abaixo de 50%."""
    if "Justificativa (obrigatório para alcance baixo)" in text:
        # Verifica % de execução para confirmar
        pct_m = re.search(r"([\d,\.]+)%\s+[\d,\.]+%\s*\nEmpenhado", text)
        if pct_m:
            try:
                pct = float(pct_m.group(1).replace(",", "."))
                if pct >= 100:   return "100% (Integral)"
                elif pct >= 75:  return "75% a 99% (Alto)"
                elif pct >= 50:  return "50% a 74% (Médio)"
                else:            return "Abaixo de 50% (Baixo)"
            except ValueError:
                pass
        return "Abaixo de 50% (Baixo)"
    # Sem justificativa: usa % de execução
    pct_m = re.search(r"([\d,\.]+)%\s+[\d,\.]+%\s*\nEmpenhado", text)
    if pct_m:
        try:
            pct = float(pct_m.group(1).replace(",", "."))
            if pct >= 100:   return "100% (Integral)"
            elif pct >= 75:  return "75% a 99% (Alto)"
            elif pct >= 50:  return "50% a 74% (Médio)"
            else:            return "Abaixo de 50% (Baixo)"
        except ValueError:
            pass
    return "100% (Integral)"


def _parse_one_meta(text):
    nm = re.search(r"META ESPECÍFICA (\d+)", text)
    if not nm:
        return None
    num = nm.group(1)

    # Descrição
    lines = text.split("\n")
    desc_lines, after = [], False
    for ln in lines:
        if re.search(r"META ESPECÍFICA \d+", ln):
            after = True
            continue
        if after:
            s = re.sub(r"\s*Total Planejado.*", "", ln.strip())
            s = re.sub(r"\s*\$[\d,\.]+\s*", "", s).strip()
            if not s or "Planejado:" in s or "Aprovada" in s:
                break
            desc_lines.append(s)
    desc_raw = " ".join(desc_lines).strip()

    plan_m = re.search(r"Planejado:\s*\$?([\d,\.]+)", text)
    exec_m = re.search(r"Executado:\s*\$?([\d,\.]+)", text)
    planejado = plan_m.group(1) if plan_m else ""
    executado = exec_m.group(1) if exec_m else ""

    emp_idx = text.find("Empenhado")
    pcts = re.findall(r"([\d,\.]+)%", text[:emp_idx]) if emp_idx > 0 else []
    exec_pct = pcts[-1].replace(".", ",") + "%" if pcts else ""

    pol_m = re.search(r"Polaridade do Indicador:[^\n]*(Quanto \w+, \w+)", text)
    polaridade = pol_m.group(1) if pol_m else ""

    ind_m = re.search(r"Indicador da Meta Específica\s+Fonte dos Dados\s*\n(.+)", text)
    if ind_m:
        raw   = ind_m.group(1).strip()
        parts = re.split(r"\s{3,}", raw)
        indicador = re.sub(r"\s+", " ", parts[0].strip())
        fonte     = re.sub(r"\s+", " ", parts[1].strip()) if len(parts) > 1 else ""
        # Separa indicador da fonte quando colados sem espaços suficientes
        for token in ["ANUÁRIO CBMERJ", "ANUÁRIOCBMERJ", "SEPOL", "Delegacia", "Departamento"]:
            if not fonte and token.rstrip() in indicador:
                idx = indicador.find(token.rstrip())
                fonte     = indicador[idx:].replace("ANUÁRIOCBMERJ", "ANUÁRIO CBMERJ").strip()
                indicador = indicador[:idx].strip()
                break
    else:
        indicador = fonte = ""

    pac_m = re.search(
        r"Pactuada no Plano.*?:\s+Valor de Referência \(Apresentado.*?\):\s*\n\s*(\S+)\s+(\S+)", text
    )
    meta_pac  = pac_m.group(1) if pac_m else ""
    val_ref_p = pac_m.group(2) if pac_m else ""

    # Layout PDF: alcance% vem antes do val_ref_monitorado
    alc_m = re.search(r"análise\):\s*\n\s*(\S+)\s*\n\s*(\S+)", text)
    if alc_m:
        val_alc     = alc_m.group(1)
        val_ref_mon = alc_m.group(2)
    else:
        val_ref_mon = val_alc = ""

    resultado = _resultado_from_block(text)

    return {
        "numero":             num,
        "polaridade":         polaridade,
        "exec_pct":           exec_pct,
        "indicador":          indicador,
        "fonte":              fonte,
        "meta_pactuada":      meta_pac,
        "val_ref_plano":      val_ref_p,
        "val_ref_monitorado": val_ref_mon,
        "val_alcance":        val_alc,
        "resultado":          resultado,
        "bens":               [],
        "_desc":              desc_raw,
        "_plan":              planejado,
        "_exec":              executado,
    }


def _extract_metas_especificas(pages):
    full = re.sub(r"Resultados dos Indicadores.*", "", "\n".join(pages), flags=re.DOTALL)
    blocos = re.split(r"(?=\bMETA ESPECÍFICA \d+\b)", full)
    metas, seen = [], set()
    for b in blocos:
        if "META ESPECÍFICA" not in b:
            continue
        b = re.sub(r"2\. Desempenho das Metas Específicas\s*\n", "", b)
        r = _parse_one_meta(b)
        if r and r["numero"] not in seen:
            seen.add(r["numero"])
            r["descricao"] = (
                f"META ESPECÍFICA {r['numero']}:  {r.pop('_desc')}"
                f"    VALOR PLANEJADO  R$ {r.pop('_plan')}"
                f"    VALOR EXECUTADO  R$ {r.pop('_exec')}"
                f"    STATUS  EM EXECUÇÃO"
            )
            metas.append(r)
    return sorted(metas, key=lambda x: int(x["numero"]))


# ─── Bens Adquiridos ───────────────────────────────────────────────────────

def _extract_bens_por_meta(pages):
    full = "\n".join(pages)
    start = full.find("6.1. Detalhamento dos Itens por Meta Específica")
    if start == -1:
        return {}
    items_text = full[start:]
    bens_por_meta = {}
    blocos = re.split(r"\nMeta Específica (\d+) —", items_text)
    i = 1
    while i < len(blocos) - 1:
        num   = blocos[i].strip()
        bloco = blocos[i + 1]
        bens  = []
        for ib in re.split(r"\n(?=\d+\. )", bloco):
            nome_m = re.match(r"(\d+\.\s+.+?)(?:\n|$)", ib)
            if not nome_m:
                continue
            pct_m = re.search(r"([\d,\.]+)\s*%", ib)
            try:
                pct = float(pct_m.group(1).replace(",", ".")) if pct_m else 0.0
            except ValueError:
                pct = 0.0
            if pct > 0:
                qtd_m = re.search(r"Qtd:\s*([\d,\.]+)", ib)
                if not qtd_m:
                    continue
                qtd  = qtd_m.group(1)
                nome = re.sub(r"^\d+\.\s*", "", nome_m.group(1))
                nome = re.split(r"\s*[-–]\s*Direcionad", nome)[0]
                nome = re.split(r"\s*[-–]\s*Unidades da", nome)[0]
                nome = _clean(nome)
                bens.append(f"- {nome}; Qtd: {qtd}")
        bens_por_meta[num] = bens
        i += 2
    return bens_por_meta


# ─── Entry point de extração ───────────────────────────────────────────────

def extract_rga_data(file_obj):
    """
    Retorna (meta_geral_dict, [meta_especifica_dict, ...]).
    Cada meta_especifica_dict tem chaves alinhadas a ME_COLS.
    """
    pages = extract_lines_from_rga_file(file_obj)
    if not pages or not any(pages):
        return None, []

    meta_geral = _extract_meta_geral(pages)
    metas      = _extract_metas_especificas(pages)
    bens       = _extract_bens_por_meta(pages)
    for m in metas:
        m["bens"] = "\n\n".join(bens.get(m["numero"], []))

    return meta_geral, metas


# ══════════════════════════════════════════════════════════════════════════════
# GERAÇÃO DO EXCEL
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
    # replica altura
    ws_dst.row_dimensions[r_dst].height = ws_src.row_dimensions[r_src].height


def _copy_merged_ranges(ws_src, r_src, ws_dst, r_dst):
    """Replica merged ranges de uma linha-fonte para uma linha-destino."""
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
    """Preenche as 4 linhas do bloco Meta Geral."""
    for r_src, r_dst in [(MG_TITLE_ROW, MG_TITLE_ROW),
                         (MG_HEADER_ROW, MG_HEADER_ROW),
                         (MG_DATA_ROW,   MG_DATA_ROW),
                         (MG_GAP_ROW,    MG_GAP_ROW)]:
        _copy_row(tmpl_ws, r_src, ws, r_dst)

    # Replica merges do cabeçalho e dados
    for r in [MG_TITLE_ROW, MG_HEADER_ROW, MG_DATA_ROW]:
        _copy_merged_ranges(tmpl_ws, r, ws, r)

    for key, col in MG_COLS.items():
        ws.cell(row=MG_DATA_ROW, column=col).value = mg.get(key, "")


def _write_meta_especifica(ws, meta, block_idx, tmpl_ws):
    """
    Escreve um bloco de Meta Específica.
    block_idx: 1 = primeira meta esp (linha 5), 2 = segunda (linha 9), etc.
    """
    base = ME_FIRST_TITLE_ROW + (block_idx - 1) * ME_BLOCK_HEIGHT
    # Linhas do template para copiar estilos (bloco da ME 2 = linhas 5-8 no template)
    tmpl_base = ME_FIRST_TITLE_ROW

    for offset in range(ME_BLOCK_HEIGHT):
        r_src = tmpl_base + offset
        r_dst = base + offset
        _copy_row(tmpl_ws, r_src, ws, r_dst)
        _copy_merged_ranges(tmpl_ws, r_src, ws, r_dst)

    # Linha de dados = base + 2
    dr = base + 2
    for key, col in ME_COLS.items():
        val = meta.get(key, "")
        ws.cell(row=dr, column=col).value = val


def generate_rga_excel_bytes(template_path: Path, meta_geral: dict, metas: list) -> bytes:
    """
    Recebe o template, os dados extraídos e retorna os bytes do Excel preenchido.
    """
    # Carrega o template como referência de estilos
    tmpl_wb = openpyxl.load_workbook(template_path)
    tmpl_ws = tmpl_wb.active

    # Cria workbook de saída
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Planilha1"

    # Copia larguras das colunas
    for col_letter, dim in tmpl_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = dim.width

    # Bloco Meta Geral
    _write_meta_geral(ws, meta_geral, tmpl_ws)

    # Blocos das Metas Específicas
    for idx, meta in enumerate(metas, start=1):
        _write_meta_especifica(ws, meta, idx, tmpl_ws)

    # Ajusta visualização
    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"
    ws.sheet_view.zoomScale = 100

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def get_missing_cells(meta_geral: dict, metas: list) -> list:
    """Retorna lista de campos em branco para alertar o usuário."""
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
