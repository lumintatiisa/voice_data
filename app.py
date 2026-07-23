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
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

PALETTE = {
    "ink": "#0B0E1A",
    "ink_soft": "#161A2C",
    "ink_line": "rgba(11,14,26,0.10)",
    "gold": "#B8912F",
    "gold_light": "#E4CB86",
    "paper": "#F6F4EE",
    "card": "#FFFFFF",
    "text": "#14151F",
    "muted": "#767A8C",
    "up_bad": "#A8402F",
    "down_good": "#2F6E5C",
}

CUSTOM_CSS = f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Inter', sans-serif;
    }}

    .stApp {{
        background-color: {PALETTE['paper']};
    }}
    #MainMenu, footer {{visibility: hidden;}}

    /* ---------- Header ---------- */
    .tfp-header {{
        background: radial-gradient(circle at 15% 20%, #1C2036 0%, {PALETTE['ink']} 55%, #05060C 100%);
        padding: 40px 44px 34px 44px;
        border-radius: 18px;
        color: #F1EFE6;
        margin-bottom: 26px;
        border-bottom: 3px solid {PALETTE['gold']};
        box-shadow: 0 18px 40px rgba(11,14,26,0.28);
    }}
    .tfp-eyebrow {{
        text-transform: uppercase;
        letter-spacing: 0.22em;
        font-size: 0.72rem;
        font-weight: 600;
        color: {PALETTE['gold_light']};
        margin-bottom: 10px;
    }}
    .tfp-header h1 {{
        font-family: 'Fraunces', serif;
        margin: 0;
        font-size: 2.15rem;
        font-weight: 600;
        letter-spacing: -0.01em;
    }}
    .tfp-header p {{
        margin: 10px 0 0 0;
        font-size: 0.95rem;
        color: rgba(241,239,230,0.72);
        max-width: 780px;
        line-height: 1.5;
    }}

    /* ---------- Section titles ---------- */
    .section-title {{
        font-family: 'Fraunces', serif;
        font-size: 1.25rem;
        font-weight: 600;
        color: {PALETTE['text']};
        margin: 8px 0 16px 0;
        display: flex;
        align-items: baseline;
        gap: 14px;
    }}
    .section-title .rule {{
        flex-grow: 1;
        height: 1px;
        background: linear-gradient(90deg, {PALETTE['ink_line']} 0%, transparent 100%);
    }}
    .section-kicker {{
        text-transform: uppercase;
        letter-spacing: 0.14em;
        font-size: 0.7rem;
        font-weight: 700;
        color: {PALETTE['gold']};
        margin-bottom: 4px;
    }}

    /* ---------- Metric cards ---------- */
    .metric-card {{
        background: {PALETTE['card']};
        border-radius: 14px;
        padding: 20px 22px 18px 22px;
        border-top: 3px solid {PALETTE['gold']};
        box-shadow: 0 3px 14px rgba(11,14,26,0.05);
        height: 100%;
    }}
    .metric-card .label {{
        font-size: 0.74rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: {PALETTE['muted']};
        font-weight: 600;
        margin-bottom: 8px;
    }}
    .metric-card .value {{
        font-family: 'Fraunces', serif;
        font-size: 2.05rem;
        font-weight: 600;
        color: {PALETTE['text']};
        line-height: 1.05;
    }}
    .metric-card .sub {{
        margin-top: 6px;
        font-size: 0.82rem;
        color: {PALETTE['muted']};
    }}
    .metric-card .delta-up {{
        color: {PALETTE['up_bad']};
        font-size: 0.82rem;
        font-weight: 600;
    }}
    .metric-card .delta-down {{
        color: {PALETTE['down_good']};
        font-size: 0.82rem;
        font-weight: 600;
    }}
    .metric-card .delta-flat {{
        color: {PALETTE['muted']};
        font-size: 0.82rem;
        font-weight: 600;
    }}

    /* ---------- Badges / pills ---------- */
    .pill {{
        display: inline-block;
        padding: 4px 13px;
        border-radius: 999px;
        font-size: 0.7rem;
        font-weight: 700;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        border: 1px solid {PALETTE['gold']};
        color: {PALETTE['ink']};
        background: rgba(184,145,47,0.08);
    }}

    /* ---------- Spotlight ranking ---------- */
    .spotlight-row {{
        display: flex;
        gap: 16px;
        margin-bottom: 22px;
    }}
    .spotlight-card {{
        flex: 1;
        background: {PALETTE['card']};
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 3px 14px rgba(11,14,26,0.05);
        position: relative;
        overflow: hidden;
    }}
    .spotlight-rank {{
        font-family: 'Fraunces', serif;
        font-size: 2.6rem;
        font-weight: 700;
        color: rgba(184,145,47,0.20);
        position: absolute;
        top: 4px;
        right: 14px;
        line-height: 1;
    }}
    .spotlight-name {{
        font-weight: 600;
        color: {PALETTE['text']};
        font-size: 0.95rem;
        margin-bottom: 4px;
        max-width: 70%;
    }}
    .spotlight-value {{
        font-family: 'Fraunces', serif;
        font-size: 1.6rem;
        font-weight: 600;
        color: {PALETTE['gold']};
    }}

    /* ---------- Data quality box ---------- */
    .data-quality-box {{
        background: #FBF2DF;
        border: 1px solid #E7CD86;
        border-radius: 12px;
        padding: 14px 18px;
        font-size: 0.85rem;
        color: #6B4F14;
        margin-bottom: 16px;
    }}

    div[data-testid="stDataFrame"] {{
        border-radius: 12px;
        overflow: hidden;
    }}

    section[data-testid="stSidebar"] {{
        background-color: #FFFFFF;
        border-right: 1px solid {PALETTE['ink_line']};
    }}
    section[data-testid="stSidebar"] h3 {{
        font-family: 'Fraunces', serif;
        font-weight: 600;
    }}

    .sidebar-label {{
        text-transform: uppercase;
        letter-spacing: 0.10em;
        font-size: 0.7rem;
        font-weight: 700;
        color: {PALETTE['muted']};
        margin: 18px 0 2px 0;
    }}

    .footer-note {{
        text-align: center;
        color: {PALETTE['muted']};
        font-size: 0.78rem;
        margin-top: 24px;
        font-style: italic;
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def section_title(kicker: str, title: str) -> None:
    st.markdown(
        f'<div class="section-kicker">{kicker}</div>'
        f'<div class="section-title">{title}<span class="rule"></span></div>',
        unsafe_allow_html=True,
    )


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
        <div class="tfp-eyebrow">Painel de Acompanhamento Comercial</div>
        <h1>Faturas em Aberto — Título Pendente (TFP)</h1>
        <p>Leitura semanal da participação de cada consultor no total de clientes com
        fatura pendente, aberta por Linha Móvel e Linha Fixa, com histórico comparado
        entre períodos.</p>
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
    with st.expander(f"Avisos de qualidade de dados ({len(issues)})", expanded=False):
        for issue in issues:
            st.markdown(f"- **{issue.file}**: {issue.message}")

periods = period_order(df)

# ---------------------------------------------------------------------------
# Sidebar — filtros globais
# ---------------------------------------------------------------------------
st.sidebar.markdown("### Filtros")

st.sidebar.markdown('<div class="sidebar-label">Período de referência</div>', unsafe_allow_html=True)
periodo_atual = st.sidebar.selectbox(
    " ", options=periods, index=len(periods) - 1, label_visibility="collapsed"
)

st.sidebar.markdown('<div class="sidebar-label">Comparar com</div>', unsafe_allow_html=True)
periodo_comparacao_opts = ["Nenhum"] + [p for p in periods if p != periodo_atual]
periodo_comparacao = st.sidebar.selectbox(
    "  ", options=periodo_comparacao_opts, index=0, label_visibility="collapsed"
)

st.sidebar.markdown('<div class="sidebar-label">Tipo de linha</div>', unsafe_allow_html=True)
tipo_visao = st.sidebar.radio(
    "   ",
    options=["Linha Móvel", "Linha Fixa", "Consolidado (Móvel + Fixa)"],
    index=2,
    label_visibility="collapsed",
)

consultores_disponiveis = sorted(df["CONSULTOR"].unique())

st.sidebar.markdown('<div class="sidebar-label">Consultor</div>', unsafe_allow_html=True)
f_consultor = st.sidebar.multiselect("    ", options=consultores_disponiveis, label_visibility="collapsed")

st.sidebar.markdown('<div class="sidebar-label">Buscar cliente</div>', unsafe_allow_html=True)
f_busca = st.sidebar.text_input("     ", placeholder="Nome ou CNPJ", label_visibility="collapsed")

st.sidebar.markdown("---")
with st.sidebar.expander("Metodologia"):
    st.markdown(
        """
**O que é o TFP?** Lista de clientes que possuem título/fatura em aberto
(inadimplência), separada por tipo de linha (Móvel/Fixa), atualizada
semanalmente.

**Como o % é calculado?** As planilhas de origem trazem apenas os clientes
**com** débito — não existe, nelas, a carteira total de clientes por
consultor. Por isso, o percentual exibido é a **participação de cada
consultor dentro do total de clientes inadimplentes** daquele tipo de linha
e período — e não uma taxa de inadimplência sobre a base completa de
clientes.

*Fórmula:* `% = (clientes únicos do consultor com fatura em aberto) ÷
(total de clientes únicos com fatura em aberto no período e tipo de linha)`

**Deduplicação:** um mesmo CNPJ pode aparecer mais de uma vez na planilha
original (múltiplos contratos/linhas). Nas métricas de "clientes", cada CNPJ
é contado uma única vez por consultor/tipo/período. A quantidade de
registros/contratos originais fica disponível na coluna `QTD_REGISTROS` do
detalhamento.

**Correção de dados aplicada:** no arquivo `TFP_SEMANA_1_E_2.xlsx`, as abas
"TFP MÓVEL" e "TFP FIXA" estavam com os rótulos trocados em relação ao
conteúdo (confirmado por comparação de CNPJ com `TFP_SEMANA_3.xlsx`). Os
rótulos foram invertidos programaticamente — ver `data_loader.py`.

**Limitação atual:** o arquivo de origem consolida as semanas 1 e 2 em uma
única aba, sem coluna de data/semana que permita separar os registros de
cada semana individualmente. Por isso esse período é tratado como um bloco
único ("Semana 1 e 2"). Assim que uma planilha trouxer as semanas 1 e 2 em
abas/colunas separadas, o comparativo semana a semana ficará automático.

**Sobre o parceiro:** o campo `PARCEIRO` continua disponível no
detalhamento e na exportação, para contexto — apenas o recorte percentual e
o filtro de busca por parceiro foram removidos desta versão do painel, a
pedido do time de negócio.
        """
    )

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
TIPO_MAP = {
    "Linha Móvel": ["MOVEL"],
    "Linha Fixa": ["FIXA"],
    "Consolidado (Móvel + Fixa)": ["MOVEL", "FIXA"],
}


def apply_filters(data: pd.DataFrame, periodo: str, tipos: list[str]) -> pd.DataFrame:
    out = data[(data["PERIODO"] == periodo) & (data["TIPO"].isin(tipos))].copy()
    if f_consultor:
        out = out[out["CONSULTOR"].isin(f_consultor)]
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
top_consultor_tab = build_share_table(cur, "CONSULTOR")
top5_share = top_consultor_tab["% DO TOTAL"].head(5).sum() if not top_consultor_tab.empty else 0


def delta_html(cur_val: int, cmp_val: int | None) -> str:
    if cmp_val is None:
        return '<span class="delta-flat">sem período de comparação</span>'
    diff = cur_val - cmp_val
    pct = (diff / cmp_val * 100) if cmp_val else 0
    if diff > 0:
        return f'<span class="delta-up">↑ +{diff} ({pct:+.1f}%) vs {periodo_comparacao}</span>'
    if diff < 0:
        return f'<span class="delta-down">↓ {diff} ({pct:+.1f}%) vs {periodo_comparacao}</span>'
    return f'<span class="delta-flat">sem variação vs {periodo_comparacao}</span>'


st.markdown(
    f'<span class="pill">{tipo_visao}</span>&nbsp;&nbsp;<span class="pill">{periodo_atual}</span>',
    unsafe_allow_html=True,
)
st.write("")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(
        f"""<div class="metric-card">
        <div class="label">Clientes com fatura em aberto</div>
        <div class="value">{total_cur:,}</div>
        <div class="sub">{delta_html(total_cur, total_cmp)}</div>
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
    top_row = top_consultor_tab.iloc[0] if not top_consultor_tab.empty else None
    valor = f"{top_row['% DO TOTAL']:.1f}%" if top_row is not None else "—"
    nome = top_row["CONSULTOR"] if top_row is not None else "—"
    st.markdown(
        f"""<div class="metric-card">
        <div class="label">Maior participação individual</div>
        <div class="value">{valor}</div>
        <div class="sub">{nome}</div>
        </div>""",
        unsafe_allow_html=True,
    )
with c4:
    st.markdown(
        f"""<div class="metric-card">
        <div class="label">Concentração — Top 5 consultores</div>
        <div class="value">{top5_share:.1f}%</div>
        <div class="sub">do total de inadimplentes está com os 5 principais</div>
        </div>""",
        unsafe_allow_html=True,
    )

st.write("")

# ---------------------------------------------------------------------------
# Participação por Consultor
# ---------------------------------------------------------------------------
section_title("Ranking", "Participação por Consultor")

top3 = top_consultor_tab.head(3).reset_index(drop=True)
if not top3.empty:
    spotlight_html = '<div class="spotlight-row">'
    ordinals = ["I", "II", "III"]
    for i, row in top3.iterrows():
        spotlight_html += f"""
        <div class="spotlight-card">
            <div class="spotlight-rank">{ordinals[i]}</div>
            <div class="spotlight-name">{row['CONSULTOR']}</div>
            <div class="spotlight-value">{row['% DO TOTAL']:.1f}%</div>
        </div>"""
    spotlight_html += "</div>"
    st.markdown(spotlight_html, unsafe_allow_html=True)

fig = px.bar(
    top_consultor_tab.sort_values("% DO TOTAL"),
    x="% DO TOTAL",
    y="CONSULTOR",
    orientation="h",
    text="% DO TOTAL",
    color="% DO TOTAL",
    color_continuous_scale=[PALETTE["ink"], PALETTE["gold"]],
)
fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside", textfont_size=12)
fig.update_layout(
    height=max(360, 28 * len(top_consultor_tab)),
    margin=dict(l=0, r=30, t=10, b=0),
    plot_bgcolor="white",
    paper_bgcolor="white",
    xaxis_title=None,
    yaxis_title=None,
    coloraxis_showscale=False,
    font=dict(family="Inter, sans-serif", color=PALETTE["text"]),
)
st.plotly_chart(fig, use_container_width=True)

# ---------------------------------------------------------------------------
# Comparativo entre semanas
# ---------------------------------------------------------------------------
section_title("Histórico", "Comparativo entre Períodos")

evol_rows = []
for p in periods:
    sub = apply_filters(df, p, tipos_sel)
    evol_rows.append({"PERIODO": p, "CLIENTES": unique_clients(sub)})
evol_df = pd.DataFrame(evol_rows)

tab_evo1, tab_evo2 = st.tabs(["Visão geral", "Por consultor"])

with tab_evo1:
    fig3 = go.Figure()
    fig3.add_trace(
        go.Scatter(
            x=evol_df["PERIODO"],
            y=evol_df["CLIENTES"],
            mode="lines+markers+text",
            text=evol_df["CLIENTES"],
            textposition="top center",
            line=dict(color=PALETTE["gold"], width=3),
            marker=dict(size=10, color=PALETTE["ink"]),
            fill="tozeroy",
            fillcolor="rgba(184,145,47,0.08)",
        )
    )
    fig3.update_layout(
        height=340,
        margin=dict(l=0, r=10, t=10, b=0),
        plot_bgcolor="white",
        paper_bgcolor="white",
        yaxis_title="Clientes com fatura em aberto",
        xaxis_title=None,
        font=dict(family="Inter, sans-serif", color=PALETTE["text"]),
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
                color_discrete_sequence=px.colors.qualitative.Prism,
            )
            fig4.update_layout(
                height=380,
                margin=dict(l=0, r=10, t=10, b=0),
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(family="Inter, sans-serif", color=PALETTE["text"]),
            )
            st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Selecione ao menos um consultor para ver a evolução.")

# ---------------------------------------------------------------------------
# Detalhamento
# ---------------------------------------------------------------------------
section_title("Base analítica", "Detalhamento de Clientes")

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
    "Baixar detalhamento filtrado (CSV)",
    data=detail.to_csv(index=False, sep=";").encode("utf-8-sig"),
    file_name=f"tfp_detalhamento_{periodo_atual.replace(' ', '_')}.csv",
    mime="text/csv",
)

st.markdown(
    '<div class="footer-note">Dashboard de Acompanhamento Comercial · TFP — '
    "dados atualizados semanalmente via /data no repositório.</div>",
    unsafe_allow_html=True,
)
