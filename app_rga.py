import base64
from pathlib import Path

import streamlit as st

from rga_engine import (
    extract_rga_data,
    extract_rga_signature,
    extract_lines_from_rga_file,
    generate_rga_excel_bytes,
    get_missing_cells,
    REQUIRED_TEMPLATE_NAME,
)

# ─── Configuração de caminhos ──────────────────────────────────────────────
BASE_DIR          = Path(__file__).resolve().parent
LOCAL_TEMPLATE    = BASE_DIR / REQUIRED_TEMPLATE_NAME
LOGO_PATH         = BASE_DIR / "Logo.png"
FAVICON_PATH      = BASE_DIR / "favicon_mj.png"


def resolve_template():
    if LOCAL_TEMPLATE.exists():
        return LOCAL_TEMPLATE, None
    return None, f"Template obrigatório não encontrado: {REQUIRED_TEMPLATE_NAME}."


# ─── Página ────────────────────────────────────────────────────────────────
page_icon = str(FAVICON_PATH) if FAVICON_PATH.exists() else "📊"
st.set_page_config(
    page_title="Gerador de Planilha RGA - FAF",
    page_icon=page_icon,
    layout="centered",
)

st.markdown(
    """
    <style>
    .header {
      display: flex;
      align-items: center;
      gap: 16px;
    }
    .header-title {
      font-size: 1.6rem !important;
      font-weight: 600;
      margin: 0;
      line-height: 1.2;
    }
    .logo-wrap {
      width: 64px;
      height: 64px;
      border-radius: 16px;
      overflow: hidden;
      border: 1px solid #e6e6e6;
      flex: 0 0 auto;
    }
    .logo-wrap img {
      width: 64px;
      height: 64px;
      object-fit: cover;
      display: block;
    }
    .brand-bar {
      display: grid;
      grid-template-columns: repeat(5, 1fr);
      height: 6px;
      border-radius: 999px;
      overflow: hidden;
      border: 1px solid #d8dbe0;
      margin-top: 0.6rem;
      margin-bottom: 14px;
      width: 699px;
    }
    .brand-bar span:nth-child(1) { background: #00b140; }
    .brand-bar span:nth-child(2) { background: #ff1b14; }
    .brand-bar span:nth-child(3) { background: #ffd200; }
    .brand-bar span:nth-child(4) { background: #1f4bff; }
    .brand-bar span:nth-child(5) { background: #ff1b14; }
    .app-subtitle {
      margin: 0 0 16px 0;
    }
    div[data-testid="stDownloadButton"] button {
      background: #217346;
      border: 1px solid #1e6a40;
      color: #ffffff;
    }
    div[data-testid="stDownloadButton"] button:hover {
      background: #1b5e38;
      border-color: #1b5e38;
      color: #ffffff;
    }
    div[data-testid="stDownloadButton"] button:active,
    div[data-testid="stDownloadButton"] button:focus,
    div[data-testid="stDownloadButton"] button:focus-visible {
      background: #1b5e38;
      border-color: #1b5e38;
      color: #ffffff;
      box-shadow: none;
      outline: none;
    }
    .blank-cells {
      border-collapse: collapse;
      width: auto;
    }
    .blank-cells th,
    .blank-cells td {
      border: 1px solid #e6e6e6;
      padding: 8px 12px;
      text-align: left;
    }
    .blank-cells th {
      background: #fafafa;
      font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─── Cabeçalho ─────────────────────────────────────────────────────────────
logo_b64 = ""
if LOGO_PATH.exists():
    try:
        logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
    except Exception:
        pass

logo_html = ""
if logo_b64:
    logo_html = f"""
      <div class="logo-wrap">
        <img src="data:image/png;base64,{logo_b64}" alt="Logo" />
      </div>
    """

st.markdown(
    f"""
    <div class="header">
      {logo_html}
      <h1 class="header-title">Gerador de Planilha RGA - FAF</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="brand-bar">
      <span></span><span></span><span></span><span></span><span></span>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<p class="app-subtitle">Faça upload do PDF do Relatório de Gestão (RGA) e gere a planilha preenchida automaticamente.</p>',
    unsafe_allow_html=True,
)

# ─── Upload ────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader("PDF do Relatório de Gestão", type=["pdf"])

if "result" not in st.session_state:
    st.session_state.result = None

# ─── Processamento ─────────────────────────────────────────────────────────
if st.button("Processar", type="primary", disabled=uploaded_file is None):
    template_path, template_error = resolve_template()
    if template_error:
        st.error(template_error)
    else:
        try:
            with st.status("Processando PDF...", expanded=True) as status:
                status.write("Lendo PDF")
                pages = extract_lines_from_rga_file(uploaded_file)

                if not pages or not any(pages):
                    status.update(label="PDF sem texto selecionável.", state="error")
                    st.error(
                        "Não foi possível extrair texto do PDF enviado. "
                        "Esse arquivo parece ser escaneado (imagem). "
                        "Envie um PDF com texto selecionável."
                    )
                    st.session_state.result = None
                    st.stop()

                status.write("Extraindo dados do RGA")
                uploaded_file.seek(0)
                meta_geral, metas = extract_rga_data(uploaded_file)

                if not meta_geral and not metas:
                    status.update(label="Nenhum dado encontrado.", state="error")
                    st.error("Nenhuma Meta encontrada no PDF do RGA.")
                    st.session_state.result = None
                    st.stop()

                status.write("Montando planilha")
                uploaded_file.seek(0)
                sig = extract_rga_signature(pages)

                excel_bytes = generate_rga_excel_bytes(
                    template_path, meta_geral, metas
                )
                missing = get_missing_cells(meta_geral, metas)

                st.session_state.result = {
                    "excel_bytes":   excel_bytes,
                    "meta_geral":    meta_geral,
                    "metas":         metas,
                    "missing":       missing,
                    "sig":           sig,
                }
                status.update(label="Processamento concluído.", state="complete")

        except Exception as exc:
            st.exception(exc)

# ─── Resultado ─────────────────────────────────────────────────────────────
result = st.session_state.result
if result:
    metas   = result["metas"]
    missing = result["missing"]
    sig     = result.get("sig", {})

    # Identificação do plano
    uf     = sig.get("uf") or "—"
    sigla  = sig.get("sigla") or "—"
    ano    = sig.get("ano") or "—"
    st.caption(f"Plano identificado: **{uf} | {sigla} | {ano}**")

    st.subheader("Resumo")
    col1, col2, col3 = st.columns(3)
    col1.metric("Metas específicas extraídas", len(metas))
    col2.metric("Meta Geral", "✓" if result["meta_geral"].get("descricao") else "✗")
    col3.metric("Campos em branco", len(missing))

    if missing:
        st.warning("Alguns campos ficaram em branco. Veja os detalhes abaixo.")

    # Nome do arquivo de saída
    uf_slug    = uf.replace(" ", "_")
    sigla_slug = sigla.replace(" ", "_")
    file_name  = f"RGA_{uf_slug}_{sigla_slug}_{ano}.xlsx"

    st.download_button(
        "⬇ Baixar Planilha",
        data=result["excel_bytes"],
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if missing:
        st.subheader("Campos em branco")
        rows_html = "".join(f"<tr><td>{m}</td></tr>" for m in missing)
        st.markdown(
            f"""
            <table class="blank-cells">
              <thead><tr><th>Campo</th></tr></thead>
              <tbody>{rows_html}</tbody>
            </table>
            """,
            unsafe_allow_html=True,
        )
