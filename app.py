"""
Calculadora de Importação — TOTVS Protheus

App Streamlit que:
  1. Recebe o PDF do espelho da NF de importação
  2. Extrai automaticamente cabeçalho, itens e totais
  3. Permite edição manual antes de gerar
  4. Gera planilha Excel com 5 abas (Inputs, Cálculo, Fórmulas Protheus, Conferência, Guia)

Deploy:
  - GitHub → Streamlit Cloud
  - Requisitos em requirements.txt
"""

from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

from core.parser import parse_pdf
from core.calculator import calcular_todos, totais_calculados
from core.excel_builder import build_excel


# ============================================================
# Configuração da página
# ============================================================
st.set_page_config(
    page_title="Calculadora de Importação — Protheus",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# Cabeçalho
# ============================================================
st.markdown(
    """
    <div style="background: linear-gradient(90deg, #1F3864 0%, #2F5496 100%);
                padding: 18px 24px; border-radius: 6px; color: white;
                margin-bottom: 18px;">
        <h1 style="margin:0; font-size: 26px;">📦 Calculadora de Importação — Protheus</h1>
        <p style="margin: 4px 0 0; opacity: 0.9; font-size: 13px;">
            Upload do espelho da NF → planilha de cálculo pronta para apoio no lançamento e no Configurador de Tributos
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# Upload do PDF
# ============================================================
col_up, col_help = st.columns([2, 1])

with col_up:
    pdf_file = st.file_uploader(
        "**1. Envie o PDF do espelho da NF de importação**",
        type=["pdf"],
        help="Espelho gerado no Protheus pela rotina de Documento de Entrada.",
    )

with col_help:
    with st.expander("ℹ️ O que esse app faz"):
        st.markdown(
            """
            - Extrai itens, alíquotas, despesas e totais do espelho da NF
            - Detecta automaticamente itens com redução de base do ICMS
            - Gera Excel com **5 abas**: Inputs, Cálculo automático, Fórmulas para o Configurador (FISA170), Conferência contra o espelho e Guia
            - Suporta múltiplos NCMs no mesmo processo (Conv. 52/91, COFINS majorada etc.)
            """
        )

if pdf_file is None:
    st.info("Aguardando o envio do PDF do espelho da NF.")
    st.stop()


# ============================================================
# Parser
# ============================================================
with st.spinner("Extraindo dados do espelho..."):
    try:
        dados = parse_pdf(pdf_file)
    except Exception as e:
        st.error(f"Erro ao processar o PDF: {e}")
        st.stop()

if not dados["itens"]:
    st.warning(
        "⚠️ Nenhum item foi identificado no PDF. Verifique se o arquivo é realmente um "
        "espelho da NF de importação do Protheus."
    )
    with st.expander("Mostrar texto extraído (debug)"):
        st.code(dados.get("_texto_bruto", "")[:3000])
    st.stop()


# ============================================================
# Cabeçalho do processo (editável)
# ============================================================
st.markdown("### 2. Identificação do processo")
cab = dados["cabecalho"]
c1, c2, c3 = st.columns(3)
with c1:
    cab["processo"] = st.text_input("Nº do Processo", cab.get("processo", ""))
with c2:
    cab["duimp"] = st.text_input("DUIMP", cab.get("duimp", ""))
with c3:
    cab["taxa_cambio"] = st.number_input(
        "Taxa câmbio USD",
        min_value=0.0,
        value=float(cab.get("taxa_cambio") or 0.0),
        step=0.0001,
        format="%.4f",
    )

c1, c2 = st.columns([2, 1])
with c1:
    cab["fornecedor"] = st.text_input("Fornecedor", cab.get("fornecedor", ""))
with c2:
    cab["modal"] = st.selectbox(
        "Modal",
        ["AÉREO", "MARÍTIMO", "RODOVIÁRIO", "FERROVIÁRIO"],
        index=["AÉREO", "MARÍTIMO", "RODOVIÁRIO", "FERROVIÁRIO"].index(
            cab.get("modal", "AÉREO")
        ),
    )


# ============================================================
# Itens (editável em data_editor)
# ============================================================
st.markdown("### 3. Itens da nota")
st.caption(
    "Os campos em laranja no Excel são editáveis; aqui você pode ajustar valores e alíquotas "
    "antes de gerar a planilha."
)

# Monta DataFrame para edição
df_itens = pd.DataFrame(
    [
        {
            "Item": it["item"],
            "Descrição": it["descricao"],
            "NCM": it["ncm"],
            "Qtd": it["qtd"],
            "CIF (R$)": it["cif"],
            "Siscomex (R$)": it["tx_sisc"],
            "Desp Ac (R$)": it["desp_ac"],
            "Alíq II": it["aliq_ii"],
            "Alíq IPI": it["aliq_ipi"],
            "Alíq PIS": it["aliq_pis"],
            "Alíq COFINS": it["aliq_cofins"],
            "ICMS cheia": 0.18,           # default; usuário ajusta
            "Carga ICMS": it["aliq_icms"],
        }
        for it in dados["itens"]
    ]
)

edited = st.data_editor(
    df_itens,
    column_config={
        "Item": st.column_config.NumberColumn(width="small", disabled=True),
        "Descrição": st.column_config.TextColumn(width="large"),
        "NCM": st.column_config.TextColumn(width="small"),
        "Qtd": st.column_config.NumberColumn(format="%.0f", width="small"),
        "CIF (R$)": st.column_config.NumberColumn(format="%.2f"),
        "Siscomex (R$)": st.column_config.NumberColumn(format="%.2f"),
        "Desp Ac (R$)": st.column_config.NumberColumn(format="%.2f"),
        "Alíq II":     st.column_config.NumberColumn(format="%.4f", help="0.20 = 20%"),
        "Alíq IPI":    st.column_config.NumberColumn(format="%.4f"),
        "Alíq PIS":    st.column_config.NumberColumn(format="%.4f"),
        "Alíq COFINS": st.column_config.NumberColumn(format="%.4f"),
        "ICMS cheia":  st.column_config.NumberColumn(format="%.4f", help="alíquota interna da UF"),
        "Carga ICMS":  st.column_config.NumberColumn(format="%.4f", help="< cheia → há redução de base"),
    },
    hide_index=True,
    use_container_width=True,
    num_rows="fixed",
)


# ============================================================
# Aplica edições de volta no dicionário
# ============================================================
for idx, row in edited.iterrows():
    if idx < len(dados["itens"]):
        it = dados["itens"][idx]
        it["descricao"]   = row["Descrição"]
        it["ncm"]         = row["NCM"]
        it["qtd"]         = float(row["Qtd"])
        it["cif"]         = float(row["CIF (R$)"])
        it["tx_sisc"]     = float(row["Siscomex (R$)"])
        it["desp_ac"]     = float(row["Desp Ac (R$)"])
        it["aliq_ii"]     = float(row["Alíq II"])
        it["aliq_ipi"]    = float(row["Alíq IPI"])
        it["aliq_pis"]    = float(row["Alíq PIS"])
        it["aliq_cofins"] = float(row["Alíq COFINS"])
        it["aliq_icms"]   = float(row["Carga ICMS"])
        it["aliq_icms_cheia"] = float(row["ICMS cheia"])


# ============================================================
# Conferência rápida em tela
# ============================================================
st.markdown("### 4. Conferência rápida")

calc = calcular_todos(dados["itens"])
tot_calc = totais_calculados(calc)
tot_esp  = dados["totais"]

def _row(label: str, esp_key: str, calc_key: str):
    e = tot_esp.get(esp_key, 0.0) or 0.0
    c = tot_calc.get(calc_key, 0.0) or 0.0
    dif = e - c
    ok = abs(dif) < 0.02
    return label, c, e, dif, ("✅ OK" if ok else "⚠️ Diferença")

linhas = [
    _row("CIF / VMLD total",       "vcif",        "cif"),
    _row("II total",               "vii_total",   "ii"),
    _row("IPI total",              "vipi",        "ipi"),
    _row("PIS-importação total",   "vpis",        "pis"),
    _row("COFINS-importação total","vcofins",     "cofins"),
    _row("Base ICMS",              "bc_icms",     "base_icms"),
    _row("ICMS destacado",         "vicms",       "icms_destacado"),
    _row("vNF",                    "vnf",         "vnf"),
]
df_conf = pd.DataFrame(linhas, columns=["Linha", "Calculado", "Espelho NF", "Dif.", "Status"])
st.dataframe(
    df_conf.style.format({"Calculado": "{:,.2f}", "Espelho NF": "{:,.2f}", "Dif.": "{:+,.2f}"}),
    hide_index=True,
    use_container_width=True,
)

n_ok = sum(1 for l in linhas if "OK" in l[4])
if n_ok == len(linhas):
    st.success(f"✅ Todos os {n_ok} totais conferem ao centavo. Pronto para gerar a planilha.")
else:
    st.warning(
        f"⚠️ {len(linhas) - n_ok} linha(s) com diferença — ajuste os inputs na tabela acima "
        "(alíquotas, Siscomex, Desp Ac) até zerar. A planilha será gerada de qualquer forma; "
        "use a aba 'Conferência' para investigar."
    )


# ============================================================
# Download da planilha
# ============================================================
st.markdown("### 5. Gerar planilha Excel")

xlsx_bytes = build_excel(dados)
filename = f"Calculo_Importacao_Proc{cab.get('processo','SEMNUM')}_{date.today().isoformat()}.xlsx"

st.download_button(
    label="⬇️ Baixar planilha de cálculo (.xlsx)",
    data=xlsx_bytes,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

st.caption(
    "A planilha tem 5 abas — Inputs, Cálculo (memória completa), Fórmulas Protheus "
    "(prontas para o Configurador), Conferência (contra o espelho) e Guia."
)


# ============================================================
# Rodapé
# ============================================================
st.divider()
st.caption(
    "Calculadora de Importação · TOTVS Protheus FISA170 · "
    "Markecode Codificação Industrial · Versão 1.0"
)
