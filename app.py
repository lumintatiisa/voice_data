# -*- coding: utf-8 -*-
"""
Dashboard de Acompanhamento Comercial — TFP (Título/Fatura Pendente)
Autor: gerado com apoio de Claude (Anthropic)

Como rodar localmente:
    streamlit run app.py

Como atualizar os dados semanalmente:
    1. Baixe o(s) arquivo(s) mais recente(s) da pasta compartilhada (OneDrive/TFP).
    2. Salve em /data seguindo o padrão de nome: TFP_SEMANA_<n>.xlsx
       (ou TFP_SEMANA_<n>_E_<m>.xlsx para períodos consolidados).
    3. Garanta que o arquivo tenha duas abas: uma de linhas MÓVEIS e outra de
       linhas FIXAS, com as colunas: CNPJ, NOME CLIENTE, CONSULTOR, PARCEIRO.
    4. Commit + push no GitHub. O Streamlit Cloud reimplanta automaticamente.
"""
from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

from data_loader import load_all_periods, period_order

# ---------------------------------------------------------------------------
# Configuração da página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Acompanhamento Comercial · TFP",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

PALETTE = {
    "primary": "#6C2BD9",
    "primary_dark": "#4A1D9E",
    "accent": "#00C2A8",
    "warn": "#FF6B6B",
    "bg_card": "#FFFFFF",
    "bg_page": "#F4F5FA",
    "text": "#1F2130",
    "muted": "#6B7280",
}

CUSTOM_CSS = f"""
<style>
    .stApp {{
        background-color: {PALETTE['bg_page']};
    }}
    #MainMenu, footer {{visibility: hidden;}}

    .tfp-header {{
        background: linear-gradient(120deg, {PALETTE['primary']} 0%, {PALETTE['primary_dark']} 100%);
        padding: 28px 32px;
        border-radius: 16px;
        color: white;
        margin-bottom: 22px;
        box-shadow: 0 8px 24px rgba(74,29,158,0.25);
    }}
    .tfp-header h1 {{
        margin: 0;
        font-size: 1.65rem;
        font-weight: 700;
    }}
    .tfp-header p {{
        margin: 6px 0 0 0;
        font-size: 0.92rem;
        opacity: 0.9;
    }}

    .metric-card {{
        background: {PALETTE['bg_card']};
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 2px 10px rgba(31,33,48,0.06);
        border: 1px solid rgba(31,33,48,0.06);
        height: 100%;
    }}
    .metric-card .label {{
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: {PALETTE['muted']};
        font-weight: 600;
        margin-bottom: 6px;
    }}
    .metric-card .value {{
        font-size: 1.9rem;
        font-weight: 800;
        color: {PALETTE['text']};
        line-height: 1.1;
    }}
    .metric-card .delta-up {{
        color: {PALETTE['warn']};
        font-size: 0.85rem;
        font-weight: 600;
    }}
    .metric-card .delta-down {{
        color: {PALETTE['accent']};
        font-size: 0.85rem;
        font-weight: 600;
    }}
    .metric-card .delta-flat {{
        color: {PALETTE['muted']};
        font-size: 0.85rem;
        font-weight: 600;
    }}

    .section-title {{
        font-size: 1.05rem;
        font-weight: 700;
        color: {PALETTE['text']};
        margin: 6px 0 10px 0;
        display: flex;
        align-items: center;
        gap: 8px;
    }}

    .badge {{
        display: inline-block;
        padding: 3px 10px;
        border-radius: 999px;
        font-size: 0.72rem;
        font-weight: 700;
        background: rgba(108,43,217,0.1);
        color: {PALETTE['primary_dark']};
        margin-left: 6px;
    }}

    .data-quality-box {{
        background: #FFF7ED;
        border: 1px solid #FCD9A8;
        border-radius: 12px;
        padding: 14px 18px;
        font-size: 0.85rem;
        color: #7A4A00;
        margin-bottom: 16px;
    }}

    div[data-testid="stDataFrame"] {{
        border-radius: 12px;
        overflow: hidden;
    }}

    section[data-testid="stSidebar"] {{
        background-color: #FFFFFF;
        border-right: 1px solid rgba(31,33,48,0.06);
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Carga de dados
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Carregando planilhas TFP...")
def get_data():
    return load_all_periods()


df, issues = get_data()

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="tfp-header">
        <h1>📊 Acompanhamento Comercial — Faturas em Aberto (TFP)</h1>
        <p>Percentual de participação de consultores e parceiros no total de clientes com título/fatura pendente,
        segmentado por Linha Móvel e Linha Fixa, com comparativo semanal.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if df.empty:
    st.error(
        "Nenhum dado carregado. Verifique se há arquivos .xlsx válidos na pasta `/data` "
        "seguindo o padrão `TFP_SEMANA_<n>.xlsx`."
    )
    st.stop()

if issues:
    with st.expander(f"⚠️ Avisos de qualidade de dados ({len(issues)})", expanded=False):
        for issue in issues:
            st.markdown(f"- **{issue.file}**: {issue.message}")

periods = period_order(df)

# ---------------------------------------------------------------------------
# Sidebar — filtros globais
# ---------------------------------------------------------------------------
st.sidebar.markdown("### 🔎 Filtros")

periodo_atual = st.sidebar.selectbox(
    "Período de referência", options=periods, index=len(periods) - 1
)
periodo_comparacao_opts = ["Nenhum"] + [p for p in periods if p != periodo_atual]
periodo_comparacao = st.sidebar.selectbox("Comparar com", options=periodo_comparacao_opts, index=0)

tipo_visao = st.sidebar.radio(
    "Tipo de linha", options=["📱 Móvel", "☎️ Fixa", "Σ Consolidado"], index=2
)

consultores_disponiveis = sorted(df["CONSULTOR"].unique())
parceiros_disponiveis = sorted(df["PARCEIRO"].unique())

f_consultor = st.sidebar.multiselect("Consultor", options=consultores_disponiveis)
f_parceiro = st.sidebar.multiselect("Parceiro", options=parceiros_disponiveis)
f_busca = st.sidebar.text_input("Buscar cliente (nome ou CNPJ)")

st.sidebar.markdown("---")
with st.sidebar.expander("ℹ️ Metodologia"):
    st.markdown(
        """
**O que é o TFP?** Lista de clientes que possuem título/fatura em aberto
(inadimplência), separada por tipo de linha (Móvel/Fixa), atualizada
semanalmente.

**Como o % é calculado?** As planilhas de origem trazem apenas os clientes
**com** débito — não existe, nelas, a carteira total de clientes por
consultor/parceiro. Por isso, o percentual exibido nos cards é a
**participação de cada consultor/parceiro dentro do total de clientes
inadimplentes** daquele tipo de linha e período — e não uma taxa de
inadimplência sobre a base completa de clientes.

*Fórmula:* `% = (clientes únicos do consultor/parceiro com fatura em aberto) ÷
(total de clientes únicos com fatura em aberto no período e tipo de linha)`

**Deduplicação:** um mesmo CNPJ pode aparecer mais de uma vez na planilha
original (múltiplos contratos/linhas). Para as métricas de "clientes", cada
CNPJ é contado uma única vez por consultor/parceiro/tipo/período. A
quantidade de registros/contratos originais fica disponível na coluna
`QTD_REGISTROS` do detalhamento.

**Correção de dados aplicada:** no arquivo `TFP_SEMANA_1_E_2.xlsx`, as abas
"TFP MÓVEL" e "TFP FIXA" estavam com os rótulos trocados em relação ao
conteúdo (confirmado por comparação de CNPJ com `TFP_SEMANA_3.xlsx`). Os
rótulos foram invertidos programaticamente — ver `data_loader.py`.

**Limitação atual:** o arquivo de origem consolida as semanas 1 e 2 em uma
única aba, sem coluna de data/semana que permita separar os registros de
cada semana individualmente. Por isso esse período é tratado como um bloco
único ("Semana 1 e 2"). Assim que uma planilha trouxer as semanas 1 e 2 em
abas/colunas separadas, o comparativo semana a semana ficará automático.
        """
    )

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
TIPO_MAP = {"📱 Móvel": ["MOVEL"], "☎️ Fixa": ["FIXA"], "Σ Consolidado": ["MOVEL", "FIXA"]}


def apply_filters(data: pd.DataFrame, periodo: str, tipos: list[str]) -> pd.DataFrame:
    out = data[(data["PERIODO"] == periodo) & (data["TIPO"].isin(tipos))].copy()
    if f_consultor:
        out = out[out["CONSULTOR"].isin(f_consultor)]
    if f_parceiro:
        out = out[out["PARCEIRO"].isin(f_parceiro)]
    if f_busca:
        needle = f_busca.strip().lower()
        out = out[
            out["NOME CLIENTE"].str.lower().str.contains(needle)
            | out["CNPJ"].str.contains(needle)
        ]
    return out


def unique_clients(data: pd.DataFrame) -> int:
    return data["CNPJ"].nunique()


def build_share_table(data: pd.DataFrame, group_col: str) -> pd.DataFrame:
    total = unique_clients(data)
    tab = (
        data.groupby(group_col)["CNPJ"]
        .nunique()
        .reset_index(name="CLIENTES_INADIMPLENTES")
        .sort_values("CLIENTES_INADIMPLENTES", ascending=False)
    )
    tab["% DO TOTAL"] = (tab["CLIENTES_INADIMPLENTES"] / total * 100).round(1) if total else 0
    return tab


tipos_sel = TIPO_MAP[tipo_visao]
cur = apply_filters(df, periodo_atual, tipos_sel)
cmp_df = (
    apply_filters(df, periodo_comparacao, tipos_sel)
    if periodo_comparacao != "Nenhum"
    else pd.DataFrame(columns=cur.columns)
)

# ---------------------------------------------------------------------------
# Cards de resumo
# ---------------------------------------------------------------------------
total_cur = unique_clients(cur)
total_cmp = unique_clients(cmp_df) if not cmp_df.empty else None
n_consultores = cur["CONSULTOR"].nunique()
n_parceiros = cur["PARCEIRO"].nunique()
top_consultor_tab = build_share_table(cur, "CONSULTOR")
top_parceiro_tab = build_share_table(cur, "PARCEIRO")


def delta_html(cur_val: int, cmp_val: int | None) -> str:
    if cmp_val is None:
        return '<span class="delta-flat">sem período de comparação</span>'
    diff = cur_val - cmp_val
    pct = (diff / cmp_val * 100) if cmp_val else 0
    if diff > 0:
        return f'<span class="delta-up">▲ +{diff} ({pct:+.1f}%) vs {periodo_comparacao}</span>'
    if diff < 0:
        return f'<span class="delta-down">▼ {diff} ({pct:+.1f}%) vs {periodo_comparacao}</span>'
    return f'<span class="delta-flat">= sem variação vs {periodo_comparacao}</span>'


st.markdown(
    f'<div class="section-title">Visão geral <span class="badge">{tipo_visao}</span>'
    f'<span class="badge">{periodo_atual}</span></div>',
    unsafe_allow_html=True,
)

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(
        f"""<div class="metric-card">
        <div class="label">Clientes com fatura em aberto</div>
        <div class="value">{total_cur:,}</div>
        {delta_html(total_cur, total_cmp)}
        </div>""",
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f"""<div class="metric-card">
        <div class="label">Consultores envolvidos</div>
        <div class="value">{n_consultores}</div>
        </div>""",
        unsafe_allow_html=True,
    )
with c3:
    st.markdown(
        f"""<div class="metric-card">
        <div class="label">Parceiros envolvidos</div>
        <div class="value">{n_parceiros}</div>
        </div>""",
        unsafe_allow_html=True,
    )
with c4:
    top_row = top_consultor_tab.iloc[0] if not top_consultor_tab.empty else None
    valor = f"{top_row['% DO TOTAL']:.1f}%" if top_row is not None else "—"
    nome = top_row["CONSULTOR"] if top_row is not None else "—"
    st.markdown(
        f"""<div class="metric-card">
        <div class="label">Maior participação (consultor)</div>
        <div class="value">{valor}</div>
        <span class="delta-flat">{nome}</span>
        </div>""",
        unsafe_allow_html=True,
    )

st.write("")

# ---------------------------------------------------------------------------
# Percentuais por Consultor / Parceiro
# ---------------------------------------------------------------------------
col_a, col_b = st.columns(2)

with col_a:
    st.markdown('<div class="section-title">👤 % por Consultor</div>', unsafe_allow_html=True)
    fig = px.bar(
        top_consultor_tab.head(20),
        x="% DO TOTAL",
        y="CONSULTOR",
        orientation="h",
        text="% DO TOTAL",
        color_discrete_sequence=[PALETTE["primary"]],
    )
    fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig.update_layout(
        yaxis={"categoryorder": "total ascending"},
        height=max(320, 26 * min(len(top_consultor_tab), 20)),
        margin=dict(l=0, r=10, t=10, b=0),
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_title=None,
        yaxis_title=None,
    )
    st.plotly_chart(fig, use_container_width=True)

with col_b:
    st.markdown('<div class="section-title">🤝 % por Parceiro</div>', unsafe_allow_html=True)
    fig2 = px.bar(
        top_parceiro_tab.head(20),
        x="% DO TOTAL",
        y="PARCEIRO",
        orientation="h",
        text="% DO TOTAL",
        color_discrete_sequence=[PALETTE["accent"]],
    )
    fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig2.update_layout(
        yaxis={"categoryorder": "total ascending"},
        height=max(320, 26 * min(len(top_parceiro_tab), 20)),
        margin=dict(l=0, r=10, t=10, b=0),
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_title=None,
        yaxis_title=None,
    )
    st.plotly_chart(fig2, use_container_width=True)

# ---------------------------------------------------------------------------
# Comparativo entre semanas
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">📈 Comparativo entre períodos</div>', unsafe_allow_html=True)

evol_rows = []
for p in periods:
    sub = apply_filters(df, p, tipos_sel)
    evol_rows.append({"PERIODO": p, "CLIENTES": unique_clients(sub)})
evol_df = pd.DataFrame(evol_rows)

tab_evo1, tab_evo2 = st.tabs(["Total geral", "Por consultor"])

with tab_evo1:
    fig3 = go.Figure()
    fig3.add_trace(
        go.Scatter(
            x=evol_df["PERIODO"],
            y=evol_df["CLIENTES"],
            mode="lines+markers+text",
            text=evol_df["CLIENTES"],
            textposition="top center",
            line=dict(color=PALETTE["primary"], width=3),
            marker=dict(size=10),
        )
    )
    fig3.update_layout(
        height=340,
        margin=dict(l=0, r=10, t=10, b=0),
        plot_bgcolor="white",
        paper_bgcolor="white",
        yaxis_title="Clientes com fatura em aberto",
        xaxis_title=None,
    )
    st.plotly_chart(fig3, use_container_width=True)

with tab_evo2:
    consultores_evo = st.multiselect(
        "Selecione consultores para comparar",
        options=consultores_disponiveis,
        default=top_consultor_tab["CONSULTOR"].head(5).tolist(),
    )
    if consultores_evo:
        rows = []
        for p in periods:
            sub = apply_filters(df, p, tipos_sel)
            sub = sub[sub["CONSULTOR"].isin(consultores_evo)]
            tab = build_share_table(apply_filters(df, p, tipos_sel), "CONSULTOR")
            tab = tab[tab["CONSULTOR"].isin(consultores_evo)]
            tab["PERIODO"] = p
            rows.append(tab)
        evo_cons = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
        if not evo_cons.empty:
            fig4 = px.line(
                evo_cons,
                x="PERIODO",
                y="% DO TOTAL",
                color="CONSULTOR",
                markers=True,
            )
            fig4.update_layout(
                height=380,
                margin=dict(l=0, r=10, t=10, b=0),
                plot_bgcolor="white",
                paper_bgcolor="white",
            )
            st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Selecione ao menos um consultor para ver a evolução.")

# ---------------------------------------------------------------------------
# Detalhamento
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">📋 Detalhamento de clientes</div>', unsafe_allow_html=True)

detail_cols = ["CNPJ", "NOME CLIENTE", "CONSULTOR", "PARCEIRO", "TIPO", "QTD_REGISTROS", "PERIODO"]
detail = cur[detail_cols].sort_values(["CONSULTOR", "NOME CLIENTE"]).reset_index(drop=True)

st.dataframe(
    detail,
    use_container_width=True,
    height=420,
    column_config={
        "CNPJ": st.column_config.TextColumn("CNPJ"),
        "QTD_REGISTROS": st.column_config.NumberColumn("Nº registros/contratos"),
    },
)

st.download_button(
    "⬇️ Baixar detalhamento filtrado (CSV)",
    data=detail.to_csv(index=False, sep=";").encode("utf-8-sig"),
    file_name=f"tfp_detalhamento_{periodo_atual.replace(' ', '_')}.csv",
    mime="text/csv",
)

st.markdown(
    f'<div style="text-align:center;color:{PALETTE["muted"]};font-size:0.78rem;margin-top:18px;">'
    f"Dashboard de Acompanhamento Comercial · TFP — dados atualizados semanalmente via /data no repositório."
    f"</div>",
    unsafe_allow_html=True,
)
