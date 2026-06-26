"""
app_streamlit.py — Validador de Folha de Pagamento — Huty Contabilidade
Interface com identidade visual Huty: verde escuro #1D3A2C, fonte Lexend
"""

import streamlit as st
import os, sys, tempfile, shutil

st.set_page_config(
    page_title="Validador de Folha — Huty",
    page_icon="📋",
    layout="wide",
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
    max-width: 1100px !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
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
            <img src="data:image/png;base64,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" alt="Huty" style="height:44px; display:block;" />
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
