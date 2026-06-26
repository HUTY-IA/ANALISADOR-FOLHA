"""
app_streamlit.py — Interface web para Validação de Folha de Pagamento
Huty Contabilidade — Streamlit Cloud
"""

import streamlit as st
import os, sys, tempfile, shutil
from pathlib import Path

# ── Configuração da página ────────────────────────────────────────
st.set_page_config(
    page_title="Validador de Folha — Huty",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ── CSS customizado ───────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #F0F4F9; }
    .main-header {
        background: linear-gradient(135deg, #1A4FA0, #2E6FD4);
        padding: 24px 32px;
        border-radius: 12px;
        margin-bottom: 24px;
        color: white;
    }
    .main-header h1 { color: white; font-size: 1.6rem; margin: 0; }
    .main-header p  { color: rgba(255,255,255,0.75); margin: 6px 0 0; font-size: 0.9rem; }
    .upload-section {
        background: white;
        border: 1px solid #DDE3EC;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 16px;
    }
    .result-ok   { background:#E2EFDA; border-left:4px solid #1A7A42; padding:12px 16px; border-radius:6px; }
    .result-warn { background:#FFF3CD; border-left:4px solid #B45309; padding:12px 16px; border-radius:6px; }
    .result-err  { background:#FFDCE0; border-left:4px solid #C62828; padding:12px 16px; border-radius:6px; }
    .metric-card {
        background: white;
        border: 1px solid #DDE3EC;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
    }
    div[data-testid="stFileUploader"] { border-radius: 8px; }
    .stButton>button {
        background: #1A4FA0;
        color: white;
        font-weight: 700;
        font-size: 1rem;
        padding: 12px 32px;
        border-radius: 8px;
        border: none;
        width: 100%;
    }
    .stButton>button:hover { background: #2E6FD4; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📋 Validador de Folha de Pagamento</h1>
    <p>Huty Contabilidade · Cruzamento automático: Apontamentos · Empréstimos · Coparticipação</p>
</div>
""", unsafe_allow_html=True)

# ── Aviso de segurança ────────────────────────────────────────────
st.info("🔒 Processamento 100% local — os arquivos são processados e descartados imediatamente. Nenhum dado é armazenado.", icon=None)

# ── Upload dos arquivos ───────────────────────────────────────────
st.markdown("### 📂 Arquivos para validação")

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
        "🏥 Coparticipação (opcional)",
        type=["xls", "xlsx"],
        help="DESCONTOS_DE_COOPARTICIPACAO_MM-AAAA.xls"
    )

# Status dos uploads
arquivos_ok = extrato and apontamentos and emprestimos
if arquivos_ok:
    st.success(f"✅ Arquivos prontos: {extrato.name} · {apontamentos.name} · {emprestimos.name}" + 
               (f" · {coop.name}" if coop else ""))
else:
    faltando = []
    if not extrato:      faltando.append("Extrato PDF")
    if not apontamentos: faltando.append("Apontamentos")
    if not emprestimos:  faltando.append("Empréstimos")
    st.warning(f"⚠️ Faltam: {', '.join(faltando)}")

st.markdown("---")

# ── Botão de execução ─────────────────────────────────────────────
if st.button("▶ Analisar e Validar", disabled=not arquivos_ok):
    
    with st.spinner("Processando arquivos..."):
        
        # Salva arquivos temporariamente
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

            # Importa e executa o engine v2
            sys.path.insert(0, '/app')
            from validador_v2 import executar

            path_relatorio = os.path.join(tmp, 'Validacao_Folha.xlsx')
            st.info("⏳ Lendo extrato PDF...")
            destino = executar(path_extrato, path_apo, path_emp, path_coop, path_relatorio)
            st.success("✅ Validação concluída!")

            # Lê o relatório gerado para mostrar resumo
            import pandas as pd
            xl = pd.ExcelFile(destino, engine='openpyxl')

            # Extrai resumo
            df_res = xl.parse('Resumo Executivo', header=None)
            resumo = {}
            for _, row in df_res.iterrows():
                k = str(row.iloc[0]).strip()
                v = str(row.iloc[1]).strip() if len(row) > 1 else ''
                if k not in ('nan', '') and v not in ('nan', ''):
                    resumo[k] = v

            empresa     = resumo.get('Empresa', '')
            competencia = resumo.get('Competência', '')
            n_func      = resumo.get('Empregados na folha', '')
            itens_ok    = resumo.get('Itens OK', '')
            divs        = resumo.get('Divergências encontradas', '0')
            atencoes    = resumo.get('Pontos de atenção', '0')
            taxa        = resumo.get('Taxa de conformidade', '')

            n_divs = int(divs) if str(divs).isdigit() else 0

            # ── Resultado visual ──────────────────────────────────
            st.markdown(f"## Resultado — {empresa} · {competencia}")

            col_a, col_b, col_c, col_d = st.columns(4)
            col_a.metric("Funcionários", n_func)
            col_b.metric("Itens OK", itens_ok)
            col_c.metric("Divergências", divs, delta=None)
            col_d.metric("Conformidade", taxa)

            if n_divs == 0:
                st.markdown('<div class="result-ok">✅ <strong>Nenhuma divergência encontrada.</strong> Folha em conformidade com os inputs.</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="result-err">⚠️ <strong>{divs} divergência(s) encontrada(s)</strong> — revisar antes de fechar a folha.</div>', unsafe_allow_html=True)

            st.markdown("")

            # Prévia das divergências
            if n_divs > 0:
                st.markdown("#### ⚠️ Divergências encontradas")
                df_divs = xl.parse('Divergências', header=1)
                df_divs = df_divs[df_divs['origem'].notna() & (df_divs['origem'] != 'nan')]
                cols_show = ['origem','cod_emp','nome_input','rubrica','descricao_input','valor_esperado','valor_folha','diferenca','status']
                cols_exist = [c for c in cols_show if c in df_divs.columns]
                st.dataframe(df_divs[cols_exist].reset_index(drop=True), use_container_width=True, height=300)

            # Botão de download
            with open(destino, 'rb') as f:
                excel_bytes = f.read()

            comp_safe = competencia.replace('/', '_')
            st.download_button(
                label="⬇ Baixar Relatório Excel Completo (11 abas)",
                data=excel_bytes,
                file_name=f"Validacao_Folha_{comp_safe}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.markdown(f"""
**Relatório inclui 11 abas:**
Resumo Executivo · Comparativo Inputs · Divergências · Pontos de Atenção · 
Empréstimos · Resumo Funcionários · Folha Rubricas · 
Input Apontamentos · Input Coparticipação · Base Empréstimos · Fonte e Controle
""")

        except Exception as e:
            st.error(f"❌ Erro durante o processamento: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

# ── Footer ────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Huty Contabilidade Ltda · Validador de Folha v2.0 · 2026")
