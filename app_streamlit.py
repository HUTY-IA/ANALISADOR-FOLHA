"""
app_streamlit.py — Validador de Folha de Pagamento — Huty Contabilidade
Interface com identidade visual Huty: verde escuro #1D3A2C, fonte Lexend
"""

import streamlit as st
import os, sys, tempfile, shutil

st.set_page_config(
    page_title="Validador de Folha — Huty",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ── Identidade visual Huty ────────────────────────────────────────
# Verde escuro: #1D3A2C | Amarelo: #D4A017 | Verde médio: #3A8C4E | Coral: #D95F3B
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Lexend', sans-serif !important;
}

/* Fundo geral */
.stApp {
    background-color: #F2F5F0;
    font-family: 'Lexend', sans-serif;
}

/* Remove padding padrão do main */
.block-container {
    padding-top: 0 !important;
    padding-bottom: 40px !important;
    max-width: 820px !important;
}

/* ── Header com logo e listras ── */
.huty-header {
    background: #1D3A2C;
    margin: -1rem -1rem 0 -1rem;
    padding: 0;
    overflow: hidden;
    position: relative;
}
.huty-header-stripes {
    height: 5px;
    background: linear-gradient(90deg,
        #1D3A2C 0%, #1D3A2C 25%,
        #D4A017 25%, #D4A017 50%,
        #3A8C4E 50%, #3A8C4E 75%,
        #D95F3B 75%, #D95F3B 100%
    );
}
.huty-header-content {
    padding: 28px 36px 24px;
    display: flex;
    align-items: center;
    gap: 20px;
}
.huty-logo-text {
    color: white;
    font-family: 'Lexend', sans-serif;
    font-weight: 800;
    font-size: 1.6rem;
    letter-spacing: 0.12em;
    line-height: 1;
}
.huty-logo-sub {
    color: rgba(255,255,255,0.55);
    font-size: 0.72rem;
    font-weight: 300;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    margin-top: 3px;
}
.huty-header-divider {
    width: 1.5px;
    height: 40px;
    background: rgba(255,255,255,0.2);
    margin: 0 4px;
}
.huty-header-title {
    color: white;
    font-family: 'Lexend', sans-serif;
}
.huty-header-title h1 {
    font-size: 1.05rem;
    font-weight: 600;
    margin: 0;
    color: white;
    letter-spacing: 0.01em;
}
.huty-header-title p {
    font-size: 0.75rem;
    color: rgba(255,255,255,0.55);
    margin: 3px 0 0;
    font-weight: 300;
    letter-spacing: 0.04em;
}

/* ── Seção de upload ── */
.section-label {
    font-family: 'Lexend', sans-serif;
    font-size: 0.68rem;
    font-weight: 600;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: #1D3A2C;
    margin: 28px 0 12px;
    display: flex;
    align-items: center;
    gap: 8px;
}
.section-label::after {
    content: '';
    flex: 1;
    height: 1px;
    background: #1D3A2C;
    opacity: 0.15;
}

/* ── Upload cards ── */
div[data-testid="stFileUploader"] {
    background: white;
    border: 1.5px solid #D9E5D9;
    border-radius: 10px;
    padding: 4px 8px 8px;
    transition: border-color 0.2s;
}
div[data-testid="stFileUploader"]:hover {
    border-color: #1D3A2C;
}
div[data-testid="stFileUploader"] label {
    font-family: 'Lexend', sans-serif !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    color: #1D3A2C !important;
}
div[data-testid="stFileUploader"] small {
    font-family: 'Lexend', sans-serif !important;
    color: #7A9A7A !important;
    font-size: 0.72rem !important;
}

/* ── Botão principal ── */
.stButton > button {
    background: #1D3A2C !important;
    color: white !important;
    font-family: 'Lexend', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    letter-spacing: 0.06em !important;
    padding: 14px 32px !important;
    border-radius: 8px !important;
    border: none !important;
    width: 100% !important;
    transition: background 0.2s !important;
    text-transform: uppercase !important;
}
.stButton > button:hover {
    background: #2A5240 !important;
}
.stButton > button:disabled {
    background: #8FB89F !important;
    color: rgba(255,255,255,0.6) !important;
}

/* ── Alertas ── */
.stAlert {
    border-radius: 8px !important;
    font-family: 'Lexend', sans-serif !important;
    font-size: 0.83rem !important;
}

/* ── Métricas ── */
div[data-testid="metric-container"] {
    background: white;
    border: 1.5px solid #D9E5D9;
    border-radius: 10px;
    padding: 14px 16px 10px;
}
div[data-testid="metric-container"] label {
    font-family: 'Lexend', sans-serif !important;
    font-size: 0.72rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    color: #7A9A7A !important;
}
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-family: 'Lexend', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1.8rem !important;
    color: #1D3A2C !important;
}

/* ── Resultado banner ── */
.result-ok {
    background: #EAF3EC;
    border-left: 4px solid #1D3A2C;
    border-radius: 0 8px 8px 0;
    padding: 14px 18px;
    font-family: 'Lexend', sans-serif;
    font-size: 0.88rem;
    color: #1D3A2C;
    font-weight: 500;
    margin: 16px 0;
}
.result-warn {
    background: #FEF9EC;
    border-left: 4px solid #D4A017;
    border-radius: 0 8px 8px 0;
    padding: 14px 18px;
    font-family: 'Lexend', sans-serif;
    font-size: 0.88rem;
    color: #7A5A00;
    font-weight: 500;
    margin: 16px 0;
}
.result-err {
    background: #FEF0EC;
    border-left: 4px solid #D95F3B;
    border-radius: 0 8px 8px 0;
    padding: 14px 18px;
    font-family: 'Lexend', sans-serif;
    font-size: 0.88rem;
    color: #7A2A10;
    font-weight: 500;
    margin: 16px 0;
}

/* ── Tabela ── */
.dataframe {
    font-family: 'Lexend', sans-serif !important;
    font-size: 0.78rem !important;
}

/* ── Download button ── */
.stDownloadButton > button {
    background: #3A8C4E !important;
    color: white !important;
    font-family: 'Lexend', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.88rem !important;
    letter-spacing: 0.05em !important;
    border-radius: 8px !important;
    border: none !important;
    width: 100% !important;
    padding: 12px !important;
}
.stDownloadButton > button:hover {
    background: #2D6B3C !important;
}

/* ── Spinner ── */
.stSpinner > div {
    border-top-color: #1D3A2C !important;
}

/* ── Segurança ── */
.security-badge {
    background: #EAF3EC;
    border: 1px solid #C5DFC8;
    border-radius: 8px;
    padding: 10px 16px;
    font-family: 'Lexend', sans-serif;
    font-size: 0.75rem;
    color: #1D3A2C;
    display: flex;
    align-items: center;
    gap: 8px;
    margin: 12px 0 4px;
}

/* Footer stripes */
.huty-footer {
    margin: 40px -1rem -1rem;
    font-family: 'Lexend', sans-serif;
}
.huty-footer-stripes {
    height: 4px;
    background: linear-gradient(90deg,
        #1D3A2C 0%, #1D3A2C 25%,
        #D4A017 25%, #D4A017 50%,
        #3A8C4E 50%, #3A8C4E 75%,
        #D95F3B 75%, #D95F3B 100%
    );
}
.huty-footer-bar {
    background: #1D3A2C;
    padding: 12px 32px;
    text-align: center;
    color: rgba(255,255,255,0.4);
    font-size: 0.72rem;
    letter-spacing: 0.1em;
    font-weight: 300;
}

/* Oculta elementos desnecessários do Streamlit */
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Header Huty ───────────────────────────────────────────────────
st.markdown("""
<div class="huty-header">
    <div class="huty-header-stripes"></div>
    <div class="huty-header-content">
        <div>
            <div class="huty-logo-text">HUTY</div>
            <div class="huty-logo-sub">Contabilidade</div>
        </div>
        <div class="huty-header-divider"></div>
        <div class="huty-header-title">
            <h1>Validador de Folha de Pagamento</h1>
            <p>Cruzamento automático · Apontamentos · Empréstimos · Coparticipação</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Segurança ─────────────────────────────────────────────────────
st.markdown("""
<div class="security-badge">
    🔒 Processamento 100% seguro — arquivos processados e descartados imediatamente. Nenhum dado é armazenado.
</div>
""", unsafe_allow_html=True)

# ── Upload ────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Arquivos para validação</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    extrato = st.file_uploader(
        "📄 Extrato Mensal PDF *",
        type=["pdf"],
        help="Aponatmentos__-_Sistema.pdf — gerado pelo sistema Domínio"
    )
    emprestimos = st.file_uploader(
        "🏦 Empréstimos *",
        type=["xls", "xlsx"],
        help="EMPRESTIMOS_MM-AAAA.xlsx"
    )
with col2:
    apontamentos = st.file_uploader(
        "📊 Apontamentos *",
        type=["xls", "xlsx"],
        help="APONTAMENTOS_MM-AAAA.xls"
    )
    coop = st.file_uploader(
        "🏥 Coparticipação",
        type=["xls", "xlsx"],
        help="DESCONTOS_DE_COOPARTICIPACAO_MM-AAAA.xls (opcional)"
    )

# Status
arquivos_ok = extrato and apontamentos and emprestimos
if arquivos_ok:
    nomes = f"{extrato.name} · {apontamentos.name} · {emprestimos.name}"
    if coop: nomes += f" · {coop.name}"
    st.success(f"✅ {nomes}")
else:
    faltando = []
    if not extrato:      faltando.append("Extrato PDF")
    if not apontamentos: faltando.append("Apontamentos")
    if not emprestimos:  faltando.append("Empréstimos")
    st.warning(f"Faltam: {', '.join(faltando)}")

st.markdown("---")

# ── Botão de análise ──────────────────────────────────────────────
if st.button("▶  ANALISAR E VALIDAR", disabled=not arquivos_ok):

    with st.spinner("Processando arquivos..."):
        tmp = tempfile.mkdtemp()
        try:
            path_extrato = os.path.join(tmp, extrato.name)
            path_apo     = os.path.join(tmp, apontamentos.name)
            path_emp     = os.path.join(tmp, emprestimos.name)
            path_coop    = os.path.join(tmp, coop.name) if coop else None

            with open(path_extrato, 'wb') as f: f.write(extrato.getvalue())
            with open(path_apo,     'wb') as f: f.write(apontamentos.getvalue())
            with open(path_emp,     'wb') as f: f.write(emprestimos.getvalue())
            if coop and path_coop:
                with open(path_coop, 'wb') as f: f.write(coop.getvalue())

            sys.path.insert(0, '/app')
            from validador_v2 import executar

            path_rel = os.path.join(tmp, 'Validacao_Folha.xlsx')
            destino  = executar(path_extrato, path_apo, path_emp, path_coop, path_rel)

            import pandas as pd
            xl = pd.ExcelFile(destino, engine='openpyxl')
            df_res = xl.parse('Resumo Executivo', header=None)

            resumo = {}
            for _, row in df_res.iterrows():
                k = str(row.iloc[0]).strip()
                v = str(row.iloc[1]).strip() if len(row) > 1 else ''
                if k not in ('nan','') and v not in ('nan',''): resumo[k] = v

            empresa     = resumo.get('Empresa', '')
            competencia = resumo.get('Competência', '')
            n_func      = resumo.get('Empregados na folha', '')
            itens_ok    = resumo.get('Itens OK', '')
            divs        = resumo.get('Divergências encontradas', '0')
            taxa        = resumo.get('Taxa de conformidade', '')
            n_divs      = int(divs) if str(divs).isdigit() else 0

            # Resultado
            st.markdown(f'<div class="section-label">{empresa} · Competência {competencia}</div>', unsafe_allow_html=True)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Funcionários", n_func)
            c2.metric("Itens verificados", resumo.get('Itens comparados contra inputs', ''))
            c3.metric("Itens OK", itens_ok)
            c4.metric("Conformidade", taxa)

            if n_divs == 0:
                st.markdown('<div class="result-ok">✅  <strong>Nenhuma divergência encontrada.</strong> Folha em conformidade com todos os inputs.</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="result-err">⚠️  <strong>{divs} divergência(s) encontrada(s).</strong> Revisar os itens abaixo antes de fechar a folha.</div>', unsafe_allow_html=True)

            # Divergências
            if n_divs > 0:
                st.markdown('<div class="section-label">Divergências</div>', unsafe_allow_html=True)
                df_divs = xl.parse('Divergências', header=1)
                df_divs = df_divs[df_divs['origem'].notna() & (df_divs['origem'] != 'nan')]
                cols_show = ['origem','cod_emp','nome_input','rubrica','descricao_input','valor_esperado','valor_folha','diferenca','status']
                cols_exist = [c for c in cols_show if c in df_divs.columns]
                st.dataframe(df_divs[cols_exist].reset_index(drop=True), use_container_width=True, height=280)

            # Empréstimos com divergência
            df_emp = xl.parse('Empréstimos', header=1)
            emp_divs = df_emp[df_emp['status'] == 'DIVERGÊNCIA'] if 'status' in df_emp.columns else pd.DataFrame()
            if len(emp_divs) > 0:
                st.markdown('<div class="section-label">Empréstimos com divergência</div>', unsafe_allow_html=True)
                cols_emp = ['nome_input','contrato','valor_input','valor_folha','diferenca','status','critica']
                cols_emp_exist = [c for c in cols_emp if c in emp_divs.columns]
                st.dataframe(emp_divs[cols_emp_exist].reset_index(drop=True), use_container_width=True, height=200)

            # Download
            st.markdown('<div class="section-label">Relatório completo</div>', unsafe_allow_html=True)
            with open(destino, 'rb') as f:
                excel_bytes = f.read()

            comp_safe = competencia.replace('/', '_')
            st.download_button(
                label="⬇  BAIXAR RELATÓRIO EXCEL — 11 ABAS",
                data=excel_bytes,
                file_name=f"Validacao_Folha_{comp_safe}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.caption("Resumo Executivo · Comparativo Inputs · Divergências · Pontos de Atenção · Empréstimos · Resumo Funcionários · Folha Rubricas · Input Apontamentos · Input Coparticipação · Base Empréstimos · Fonte e Controle")

        except Exception as e:
            st.error(f"❌ Erro: {str(e)}")
            import traceback
            with st.expander("Detalhes do erro"):
                st.code(traceback.format_exc())
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

# ── Footer ────────────────────────────────────────────────────────
st.markdown("""
<div class="huty-footer">
    <div class="huty-footer-stripes"></div>
    <div class="huty-footer-bar">HUTY CONTABILIDADE LTDA · VALIDADOR DE FOLHA V3.0 · 2026</div>
</div>
""", unsafe_allow_html=True)
