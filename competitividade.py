import streamlit as st
import pandas as pd
from textwrap import dedent
from html import escape

st.set_page_config(page_title="Competitividade Brasil", layout="wide")

NAVY = "#011D35"
YELLOW = "#F7C223"
GREEN = "#00823B"
RED = "#C81E36"
BLUE = "#0B5EA8"
LIGHT_BG = "#F4F7FB"
BORDER = "#D7E0EA"
TEXT = "#16324F"
MUTED = "#6A7A8C"

ALL_FACTORS = [
    "Ambiente de Negócios",
    "Ambiente Econômico",
    "Baixo Carbono e Recursos Naturais",
    "Comércio e Integração Internacional",
    "Desenvolvimento Humano e Trabalho",
    "Desenvolvimento Produtivo, Inovação e Tecnologia",
    "Educação",
    "Infraestrutura",
]

ARQUIVO_EXCEL = "dados_dashboard_2025 (6).xlsx"


def render_html(raw_html: str):
    raw_html = dedent(str(raw_html)).strip()
    if hasattr(st, "html"):
        st.html(raw_html)
    else:
        raw_html = "\n".join(line.lstrip() for line in raw_html.splitlines())
        st.markdown(raw_html, unsafe_allow_html=True)


render_html(
    f"""
    <style>
    .stApp {{
        background-color: {LIGHT_BG};
    }}

    .block-container {{
        padding-top: 3.5rem;
        padding-bottom: 2rem;
        max-width: 1500px;
    }}

    .hero {{
        margin-top: 0.8rem;
        overflow: visible;
        background: linear-gradient(135deg, #03294a 0%, {NAVY} 75%);
        border-radius: 28px;
        padding: 42px 34px 32px 34px;
        color: white;
        margin-bottom: 1rem;
        box-shadow: 0 10px 24px rgba(1,29,53,0.16);
    }}

    .hero-kicker {{
        display:inline-block;
        padding:6px 12px;
        border-radius:999px;
        background:rgba(255,255,255,.10);
        color:{YELLOW};
        font-size:12px;
        font-weight:800;
        letter-spacing:.06em;
        text-transform:uppercase;
        margin-bottom:10px;
    }}

    .hero-title {{
        font-size: 2.15rem;
        font-weight: 850;
        line-height: 1.18;
        margin: 0 0 8px 0;
    }}

    .hero-sub {{
        color:#E6EDF5;
        font-size:1rem;
    }}

    .metric-card {{
        background:white;
        border:1px solid {BORDER};
        border-radius:22px;
        padding:18px;
        box-shadow: 0 4px 12px rgba(1,29,53,.04);
        height:100%;
    }}

    .metric-num {{
        font-size:2.1rem;
        font-weight:850;
        line-height:1;
    }}

    .metric-lab {{
        margin-top:6px;
        color:{TEXT};
        font-weight:750;
        font-size:0.98rem;
    }}

    .metric-sub {{
        margin-top:4px;
        color:{MUTED};
        font-size:0.82rem;
    }}

    .factor-main {{
        background:white;
        border:1px solid {BORDER};
        border-radius:24px;
        padding:22px;
        box-shadow: 0 5px 14px rgba(1,29,53,.05);
        margin-bottom: 1rem;
        border-left: 8px solid {YELLOW};
    }}

    .eyebrow {{
        font-size:11px;
        letter-spacing:.16em;
        text-transform:uppercase;
        color:{MUTED};
        font-weight:850;
        margin-bottom: 8px;
    }}

    .factor-title {{
        font-size:2rem;
        font-weight:850;
        color:{NAVY};
        margin-bottom: 10px;
        line-height:1.15;
    }}

    .rank-line {{
        font-size:2.5rem;
        font-weight:850;
        color:{NAVY};
        line-height:1.05;
    }}

    .score-line {{
        color:{MUTED};
        font-size:.9rem;
        margin-top:8px;
    }}

    .badge {{
        display:inline-flex;
        align-items:center;
        gap:6px;
        padding:7px 12px;
        border-radius:999px;
        font-size:12px;
        font-weight:800;
        border:1px solid transparent;
        margin-top:12px;
        white-space: nowrap;
    }}

    .badge-good {{
        background:#E8F6EC;
        border-color:#CDECD6;
        color:{GREEN};
    }}

    .badge-bad {{
        background:#FCEBED;
        border-color:#F5C8D0;
        color:{RED};
    }}

    .badge-neutral {{
        background:#F1F5F9;
        border-color:#D9E2EC;
        color:#536579;
    }}

    .reading-box {{
        background:{NAVY};
        color:white;
        border-radius:20px;
        padding:16px 18px;
        margin: 0.4rem 0 1.1rem 0;
        font-size:0.98rem;
        line-height:1.5;
    }}

    .reading-box b {{
        color:white;
    }}

    .factor-overview-card {{
        background:white;
        border:1px solid {BORDER};
        border-radius:22px;
        padding:18px;
        box-shadow: 0 4px 12px rgba(1,29,53,.04);
        margin-bottom: 0.8rem;
        border-left: 6px solid {YELLOW};
    }}

    .factor-overview-name {{
        font-size:1.12rem;
        font-weight:850;
        color:{NAVY};
        line-height:1.25;
    }}

    .factor-overview-rank {{
        font-size:1.9rem;
        font-weight:850;
        color:{NAVY};
        margin-top:10px;
        line-height:1;
    }}

    .small-muted {{
        color:{MUTED};
        font-size:.82rem;
        margin-top:6px;
    }}

    .subfactor-wrap {{
        position:relative;
        margin: 0 0 1.6rem 0;
    }}

    .subfactor-card {{
        background:white;
        border:1px solid {BORDER};
        border-radius:24px;
        padding:20px 22px;
        box-shadow: 0 8px 18px rgba(1,29,53,.05);
        border-left: 8px solid {YELLOW};
        position: relative;
        z-index:2;
    }}

    .subfactor-head {{
        display:flex;
        justify-content:space-between;
        align-items:flex-start;
        gap:16px;
    }}

    .subfactor-name {{
        font-size:1.55rem;
        font-weight:850;
        color:{NAVY};
        line-height:1.15;
        margin-top:4px;
    }}

    .subfactor-rank {{
        font-size:2.15rem;
        font-weight:850;
        color:{NAVY};
        line-height:1.05;
        margin-top:12px;
    }}

    .subfactor-score {{
        color:{MUTED};
        font-size:.84rem;
        margin-top:6px;
    }}

    .children-area {{
        position:relative;
        padding:20px 0 0 34px;
        margin-left: 22px;
    }}

    .children-area::before {{
        content:"";
        position:absolute;
        top:0;
        left:13px;
        width:2px;
        height:calc(100% - 16px);
        background:#BFCBDC;
        border-radius:999px;
    }}

    .indicator-grid {{
        display:grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap:14px;
    }}

    .indicator-card {{
        position:relative;
        background:white;
        border:1px solid #DFE7F0;
        border-radius:20px;
        padding:16px 16px 14px 42px;
        min-height:150px;
        box-shadow: 0 4px 10px rgba(1,29,53,.035);
    }}

    .indicator-card::before {{
        content:"";
        position:absolute;
        left:-22px;
        top:30px;
        width:22px;
        height:2px;
        background:#BFCBDC;
    }}

    .indicator-card::after {{
        content:"";
        position:absolute;
        left:14px;
        top:18px;
        width:14px;
        height:14px;
        border-radius:50%;
        background:{YELLOW};
        box-shadow: 0 0 0 4px rgba(247,194,35,.18);
    }}

    .indicator-type {{
        font-size:10px;
        letter-spacing:.14em;
        text-transform:uppercase;
        color:{MUTED};
        font-weight:900;
        margin-bottom:6px;
    }}

    .indicator-name {{
        font-size:1.03rem;
        font-weight:850;
        color:{NAVY};
        line-height:1.25;
        margin-bottom:12px;
        min-height:42px;
    }}

    .indicator-rank {{
        font-size:1.7rem;
        font-weight:850;
        color:{NAVY};
        line-height:1.05;
    }}

    .indicator-score {{
        color:{MUTED};
        font-size:.8rem;
        margin-top:6px;
        margin-bottom:8px;
    }}

    .yellow-arrow {{
        color:{YELLOW};
        padding:0 6px;
    }}

    @media (max-width: 900px) {{
        .indicator-grid {{
            grid-template-columns: 1fr;
        }}
        .hero-title {{
            font-size: 1.8rem;
        }}
        .rank-line {{
            font-size:2rem;
        }}
        .subfactor-head {{
            flex-direction:column;
        }}
    }}
    </style>
    """
)


def normalize_columns(df):
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )
    return df


def fmt_rank(x):
    if pd.isna(x):
        return "-"
    try:
        return f"{int(round(float(x)))}º"
    except Exception:
        return "-"


def fmt_score(x):
    if pd.isna(x):
        return "-"
    try:
        return f"{float(x):.2f}".replace(".", ",")
    except Exception:
        return "-"


def status_from_delta(delta):
    if pd.isna(delta):
        return "Estável"
    if delta > 0:
        return "Melhorou"
    if delta < 0:
        return "Piorou"
    return "Estável"


def badge_html(status, delta):
    if status == "Melhorou":
        cls = "badge badge-good"
        label = f"↗ Melhorou (+{int(delta)})"
    elif status == "Piorou":
        cls = "badge badge-bad"
        label = f"↘ Piorou ({int(delta)})"
    else:
        cls = "badge badge-neutral"
        label = "— Estável"
    return f"<span class='{cls}'>{label}</span>"


@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df = normalize_columns(df)

    for col in ["pais", "fator", "subfator", "indicador", "edicao_referencia"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    for col in ["ranking", "valor_indicador_normalizado"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def compare_item(df, pais, ed0, ed1, fator=None, subfator=None, indicador=None):
    temp = df[df["pais"] == pais].copy()

    if fator is not None:
        temp = temp[temp["fator"] == fator]
    if subfator is not None:
        temp = temp[temp["subfator"] == subfator]
    if indicador is not None:
        temp = temp[temp["indicador"] == indicador]

    r0 = temp[temp["edicao_referencia"] == ed0]
    r1 = temp[temp["edicao_referencia"] == ed1]

    if r0.empty or r1.empty:
        return None

    rank0 = r0["ranking"].iloc[0]
    rank1 = r1["ranking"].iloc[0]
    score0 = r0["valor_indicador_normalizado"].iloc[0]
    score1 = r1["valor_indicador_normalizado"].iloc[0]

    delta_rank = rank0 - rank1
    status = status_from_delta(delta_rank)

    return {
        "rank0": rank0,
        "rank1": rank1,
        "score0": score0,
        "score1": score1,
        "delta_rank": delta_rank,
        "status": status,
    }


def get_factor_summary(df, pais, ed0, ed1, fator):
    return compare_item(df, pais, ed0, ed1, fator=fator, subfator="-", indicador="-")


def get_subfactor_list(df, pais, ed0, ed1, fator):
    subfactors = (
        df[
            (df["pais"] == pais)
            & (df["fator"] == fator)
            & (df["subfator"] != "-")
            & (df["indicador"] == "-")
        ]["subfator"]
        .dropna()
        .unique()
        .tolist()
    )

    out = []
    for sf in subfactors:
        comp = compare_item(
            df, pais, ed0, ed1, fator=fator, subfator=sf, indicador="-"
        )
        if comp is not None:
            out.append({"subfator": sf, **comp})

    return out


def get_indicator_list(df, pais, ed0, ed1, fator, subfator):
    indicators = (
        df[
            (df["pais"] == pais)
            & (df["fator"] == fator)
            & (df["subfator"] == subfator)
            & (df["indicador"] != "-")
        ]["indicador"]
        .dropna()
        .unique()
        .tolist()
    )

    out = []
    for ind in indicators:
        comp = compare_item(
            df, pais, ed0, ed1, fator=fator, subfator=subfator, indicador=ind
        )
        if comp is not None:
            out.append({"indicador": ind, **comp})

    return out


def render_metric_card(value, title, subtitle, color):
    render_html(
        f"""
        <div class="metric-card">
            <div class="metric-num" style="color:{color};">{value}</div>
            <div class="metric-lab">{escape(str(title))}</div>
            <div class="metric-sub">{escape(str(subtitle))}</div>
        </div>
        """
    )


def render_factor_main_card(fator, comp):
    render_html(
        f"""
        <div class="factor-main">
            <div class="eyebrow">Fator-chave</div>
            <div class="factor-title">{escape(str(fator))}</div>
            <div class="rank-line">
                {fmt_rank(comp['rank0'])}<span class="yellow-arrow">→</span>{fmt_rank(comp['rank1'])}
            </div>
            <div class="score-line">score: {fmt_score(comp['score0'])} → {fmt_score(comp['score1'])}</div>
            {badge_html(comp['status'], comp['delta_rank'])}
        </div>
        """
    )


def render_factor_overview_card(fator, comp):
    render_html(
        f"""
        <div class="factor-overview-card">
            <div class="factor-overview-name">{escape(str(fator))}</div>
            <div class="factor-overview-rank">
                {fmt_rank(comp['rank0'])}<span class="yellow-arrow">→</span>{fmt_rank(comp['rank1'])}
            </div>
            <div class="small-muted">score: {fmt_score(comp['score0'])} → {fmt_score(comp['score1'])}</div>
            {badge_html(comp['status'], comp['delta_rank'])}
        </div>
        """
    )


def indicator_card_html(ind):
    return f"""
    <div class="indicator-card">
        <div class="indicator-type">Indicador relacionado</div>
        <div class="indicator-name">{escape(str(ind['indicador']))}</div>
        <div class="indicator-rank">
            {fmt_rank(ind['rank0'])}<span class="yellow-arrow">→</span>{fmt_rank(ind['rank1'])}
        </div>
        <div class="indicator-score">score: {fmt_score(ind['score0'])} → {fmt_score(ind['score1'])}</div>
        {badge_html(ind['status'], ind['delta_rank'])}
    </div>
    """


def render_subfactor_block(sub, indicators):
    indicators_html = "".join(indicator_card_html(ind) for ind in indicators)

    if indicators_html == "":
        indicators_html = """
        <div class="indicator-card">
            <div class="indicator-type">Indicador relacionado</div>
            <div class="indicator-name">Sem indicadores disponíveis</div>
        </div>
        """

    render_html(
        f"""
        <div class="subfactor-wrap">
            <div class="subfactor-card">
                <div class="subfactor-head">
                    <div>
                        <div class="eyebrow">Subfator</div>
                        <div class="subfactor-name">{escape(str(sub['subfator']))}</div>
                        <div class="subfactor-rank">
                            {fmt_rank(sub['rank0'])}<span class="yellow-arrow">→</span>{fmt_rank(sub['rank1'])}
                        </div>
                        <div class="subfactor-score">score: {fmt_score(sub['score0'])} → {fmt_score(sub['score1'])}</div>
                    </div>
                    <div>{badge_html(sub['status'], sub['delta_rank'])}</div>
                </div>
            </div>

            <div class="children-area">
                <div class="indicator-grid">
                    {indicators_html}
                </div>
            </div>
        </div>
        """
    )



def build_static_factor_html(pais, ed0, ed1, fator, comp, subfactors, indicators_by_subfactor):
    def clean(x):
        return escape(str(x))

    def factor_card_html():
        return f"""
        <div class="factor-main">
            <div class="eyebrow">Fator-chave</div>
            <div class="factor-title">{clean(fator)}</div>
            <div class="rank-line">
                {fmt_rank(comp['rank0'])}<span class="yellow-arrow">→</span>{fmt_rank(comp['rank1'])}
            </div>
            <div class="score-line">score: {fmt_score(comp['score0'])} → {fmt_score(comp['score1'])}</div>
            {badge_html(comp['status'], comp['delta_rank'])}
        </div>
        """

    def indicator_static_html(ind):
        return f"""
        <div class="indicator-card">
            <div class="indicator-type">Indicador relacionado</div>
            <div class="indicator-name">{clean(ind['indicador'])}</div>
            <div class="indicator-rank">
                {fmt_rank(ind['rank0'])}<span class="yellow-arrow">→</span>{fmt_rank(ind['rank1'])}
            </div>
            <div class="indicator-score">score: {fmt_score(ind['score0'])} → {fmt_score(ind['score1'])}</div>
            {badge_html(ind['status'], ind['delta_rank'])}
        </div>
        """

    def subfactor_static_html(sub):
        inds = indicators_by_subfactor.get(sub["subfator"], [])
        indicators_html = "".join(indicator_static_html(ind) for ind in inds)
        if indicators_html == "":
            indicators_html = """
            <div class="indicator-card">
                <div class="indicator-type">Indicador relacionado</div>
                <div class="indicator-name">Sem indicadores disponíveis</div>
            </div>
            """

        return f"""
        <div class="subfactor-wrap">
            <div class="subfactor-card">
                <div class="subfactor-head">
                    <div>
                        <div class="eyebrow">Subfator</div>
                        <div class="subfactor-name">{clean(sub['subfator'])}</div>
                        <div class="subfactor-rank">
                            {fmt_rank(sub['rank0'])}<span class="yellow-arrow">→</span>{fmt_rank(sub['rank1'])}
                        </div>
                        <div class="subfactor-score">score: {fmt_score(sub['score0'])} → {fmt_score(sub['score1'])}</div>
                    </div>
                    <div>{badge_html(sub['status'], sub['delta_rank'])}</div>
                </div>
            </div>

            <div class="children-area">
                <div class="indicator-grid">
                    {indicators_html}
                </div>
            </div>
        </div>
        """

    subfactor_html = "".join(subfactor_static_html(sf) for sf in subfactors)

    return f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{clean(fator)} - Competitividade Brasil</title>
<style>
:root {{
    --navy: {NAVY};
    --yellow: {YELLOW};
    --green: {GREEN};
    --red: {RED};
    --blue: {BLUE};
    --light-bg: {LIGHT_BG};
    --border: {BORDER};
    --text: {TEXT};
    --muted: {MUTED};
}}
* {{
    box-sizing: border-box;
}}
body {{
    margin: 0;
    font-family: "Segoe UI", Arial, sans-serif;
    background: var(--light-bg);
    color: var(--text);
}}
.page {{
    max-width: 1400px;
    margin: 34px auto;
    padding: 0 28px 36px 28px;
}}
.hero {{
    background: linear-gradient(135deg, #03294a 0%, var(--navy) 75%);
    border-radius: 28px;
    padding: 40px 34px 32px 34px;
    color: white;
    margin-bottom: 20px;
    box-shadow: 0 10px 24px rgba(1,29,53,0.16);
}}
.hero-kicker {{
    display:inline-block;
    padding:6px 12px;
    border-radius:999px;
    background:rgba(255,255,255,.10);
    color:var(--yellow);
    font-size:12px;
    font-weight:800;
    letter-spacing:.06em;
    text-transform:uppercase;
    margin-bottom:10px;
}}
.hero-title {{
    font-size: 2.15rem;
    font-weight: 850;
    line-height: 1.18;
    margin: 0 0 8px 0;
}}
.hero-sub {{
    color:#E6EDF5;
    font-size:1rem;
}}
.factor-main {{
    background:white;
    border:1px solid var(--border);
    border-radius:24px;
    padding:22px;
    box-shadow: 0 5px 14px rgba(1,29,53,.05);
    margin-bottom: 18px;
    border-left: 8px solid var(--yellow);
}}
.eyebrow {{
    font-size:11px;
    letter-spacing:.16em;
    text-transform:uppercase;
    color:var(--muted);
    font-weight:850;
    margin-bottom: 8px;
}}
.factor-title {{
    font-size:2rem;
    font-weight:850;
    color:var(--navy);
    margin-bottom: 10px;
    line-height:1.15;
}}
.rank-line {{
    font-size:2.5rem;
    font-weight:850;
    color:var(--navy);
    line-height:1.05;
}}
.score-line {{
    color:var(--muted);
    font-size:.9rem;
    margin-top:8px;
}}
.badge {{
    display:inline-flex;
    align-items:center;
    gap:6px;
    padding:7px 12px;
    border-radius:999px;
    font-size:12px;
    font-weight:800;
    border:1px solid transparent;
    margin-top:12px;
    white-space: nowrap;
}}
.badge-good {{
    background:#E8F6EC;
    border-color:#CDECD6;
    color:var(--green);
}}
.badge-bad {{
    background:#FCEBED;
    border-color:#F5C8D0;
    color:var(--red);
}}
.badge-neutral {{
    background:#F1F5F9;
    border-color:#D9E2EC;
    color:#536579;
}}
.reading-box {{
    background:var(--navy);
    color:white;
    border-radius:20px;
    padding:16px 18px;
    margin: 0.4rem 0 1.1rem 0;
    font-size:0.98rem;
    line-height:1.5;
}}
.subfactor-wrap {{
    position:relative;
    margin: 0 0 1.6rem 0;
}}
.subfactor-card {{
    background:white;
    border:1px solid var(--border);
    border-radius:24px;
    padding:20px 22px;
    box-shadow: 0 8px 18px rgba(1,29,53,.05);
    border-left: 8px solid var(--yellow);
    position: relative;
    z-index:2;
}}
.subfactor-head {{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:16px;
}}
.subfactor-name {{
    font-size:1.55rem;
    font-weight:850;
    color:var(--navy);
    line-height:1.15;
    margin-top:4px;
}}
.subfactor-rank {{
    font-size:2.15rem;
    font-weight:850;
    color:var(--navy);
    line-height:1.05;
    margin-top:12px;
}}
.subfactor-score {{
    color:var(--muted);
    font-size:.84rem;
    margin-top:6px;
}}
.children-area {{
    position:relative;
    padding:20px 0 0 34px;
    margin-left: 22px;
}}
.children-area::before {{
    content:"";
    position:absolute;
    top:0;
    left:13px;
    width:2px;
    height:calc(100% - 16px);
    background:#BFCBDC;
    border-radius:999px;
}}
.indicator-grid {{
    display:grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap:14px;
}}
.indicator-card {{
    position:relative;
    background:white;
    border:1px solid #DFE7F0;
    border-radius:20px;
    padding:16px 16px 14px 42px;
    min-height:150px;
    box-shadow: 0 4px 10px rgba(1,29,53,.035);
}}
.indicator-card::before {{
    content:"";
    position:absolute;
    left:-22px;
    top:30px;
    width:22px;
    height:2px;
    background:#BFCBDC;
}}
.indicator-card::after {{
    content:"";
    position:absolute;
    left:14px;
    top:18px;
    width:14px;
    height:14px;
    border-radius:50%;
    background:var(--yellow);
    box-shadow: 0 0 0 4px rgba(247,194,35,.18);
}}
.indicator-type {{
    font-size:10px;
    letter-spacing:.14em;
    text-transform:uppercase;
    color:var(--muted);
    font-weight:900;
    margin-bottom:6px;
}}
.indicator-name {{
    font-size:1.03rem;
    font-weight:850;
    color:var(--navy);
    line-height:1.25;
    margin-bottom:12px;
    min-height:42px;
}}
.indicator-rank {{
    font-size:1.7rem;
    font-weight:850;
    color:var(--navy);
    line-height:1.05;
}}
.indicator-score {{
    color:var(--muted);
    font-size:.8rem;
    margin-top:6px;
    margin-bottom:8px;
}}
.yellow-arrow {{
    color:var(--yellow);
    padding:0 6px;
}}
@media print {{
    body {{
        background:white;
    }}
    .page {{
        max-width: none;
        margin: 0;
    }}
    .subfactor-wrap {{
        break-inside: avoid;
        page-break-inside: avoid;
    }}
}}
@media (max-width: 900px) {{
    .indicator-grid {{
        grid-template-columns: 1fr;
    }}
    .hero-title {{
        font-size: 1.8rem;
    }}
    .subfactor-head {{
        flex-direction:column;
    }}
}}
</style>
</head>
<body>
<div class="page">
    <div class="hero">
        <div class="hero-kicker">Painel estático • {clean(pais)}</div>
        <div class="hero-title">{clean(fator)}: mudanças por subfator e indicador</div>
        <div class="hero-sub">Comparação entre {clean(ed0)} e {clean(ed1)}. Ranking menor representa melhor posição.</div>
    </div>

    {factor_card_html()}

    <div class="reading-box">
        <b>Leitura encadeada:</b> cada balão principal representa um subfator. Abaixo dele aparecem os indicadores relacionados, conectados visualmente por linhas e marcadores.
    </div>

    {subfactor_html}
</div>
</body>
</html>"""


def make_download_filename(fator, pais, ed0, ed1):
    safe = (
        f"{pais}_{fator}_{ed0}_{ed1}"
        .replace("/", "-")
        .replace("\\", "-")
        .replace(" ", "_")
        .replace(",", "")
        .replace(":", "")
    )
    return f"painel_estatico_{safe}.html"


# SIDEBAR
st.sidebar.title("Competitividade Brasil")
st.sidebar.caption("Dashboard encadeado: fator → subfator → indicador")

try:
    df = load_data(ARQUIVO_EXCEL)
except FileNotFoundError:
    st.error(
        "Não encontrei a planilha no caminho definido no código:\n\n"
        f"{ARQUIVO_EXCEL}\n\n"
        "Confira se o arquivo existe exatamente com esse nome."
    )
    st.stop()
except Exception as e:
    st.error(f"Erro ao ler a planilha: {e}")
    st.stop()


required_cols = {
    "pais",
    "fator",
    "subfator",
    "indicador",
    "edicao_referencia",
    "ranking",
    "valor_indicador_normalizado",
}
missing = [col for col in required_cols if col not in df.columns]

if missing:
    st.error(f"Faltam colunas obrigatórias na planilha: {', '.join(missing)}")
    st.stop()

paises = sorted(df["pais"].dropna().unique().tolist())
default_country = paises.index("Brasil") if "Brasil" in paises else 0
pais = st.sidebar.selectbox("País", paises, index=default_country)

edicoes = sorted(df[df["pais"] == pais]["edicao_referencia"].dropna().unique().tolist())

if len(edicoes) < 2:
    st.error("A base precisa ter pelo menos duas edições para comparar.")
    st.stop()

ed0 = st.sidebar.selectbox("Edição inicial", edicoes, index=0)
ed1 = st.sidebar.selectbox("Edição final", edicoes, index=len(edicoes) - 1)

if ed0 == ed1:
    st.warning("Escolha duas edições diferentes.")
    st.stop()

available_factors = []
for fator in ALL_FACTORS:
    comp = get_factor_summary(df, pais, ed0, ed1, fator)
    if comp is not None:
        available_factors.append(fator)

if not available_factors:
    st.error("Não encontrei fatores compatíveis com os filtros escolhidos.")
    st.stop()


# HEADER

render_html(
    f"""
    <div class="hero">
        <div class="hero-kicker">Dashboard único • {escape(str(pais))}</div>
        <div class="hero-title">Competitividade Brasil: mudanças por fator, subfator e indicador</div>
        <div class="hero-sub">
            Comparação entre {escape(str(ed0))} e {escape(str(ed1))}. O status é definido pela variação no ranking:
            posição menor é melhor.
        </div>
    </div>
    """
)

page_options = ["Visão Geral"] + available_factors
selected_page = st.radio(
    "Selecione a página",
    page_options,
    horizontal=True,
    label_visibility="collapsed",
)


# VISÃO GERAL

if selected_page == "Visão Geral":
    factor_comps = []

    for fator in available_factors:
        comp = get_factor_summary(df, pais, ed0, ed1, fator)
        if comp is not None:
            factor_comps.append({"fator": fator, **comp})

    n_better = sum(1 for item in factor_comps if item["status"] == "Melhorou")
    n_worse = sum(1 for item in factor_comps if item["status"] == "Piorou")
    n_stable = sum(1 for item in factor_comps if item["status"] == "Estável")

    c1, c2, c3 = st.columns(3)
    with c1:
        render_metric_card(n_better, "fatores melhoraram", "nível agregado", GREEN)
    with c2:
        render_metric_card(n_worse, "fatores pioraram", "nível agregado", RED)
    with c3:
        render_metric_card(n_stable, "fatores estáveis", "nível agregado", BLUE)

    render_html(
        f"""
        <div class="reading-box">
            <b>Leitura executiva:</b> abaixo estão os 8 fatores-chave comparados entre
            {escape(str(ed0))} e {escape(str(ed1))}. Use o seletor acima para ver o encadeamento completo de cada fator,
            com subfatores e indicadores.
        </div>
        """
    )

    for i in range(0, len(factor_comps), 2):
        cols = st.columns(2)
        with cols[0]:
            render_factor_overview_card(factor_comps[i]["fator"], factor_comps[i])
        if i + 1 < len(factor_comps):
            with cols[1]:
                render_factor_overview_card(factor_comps[i + 1]["fator"], factor_comps[i + 1])

else:
    fator = selected_page
    comp = get_factor_summary(df, pais, ed0, ed1, fator)

    if comp is None:
        st.info("Sem dados para este fator.")
        st.stop()

    st.subheader(fator)
    render_factor_main_card(fator, comp)

    subfactors = get_subfactor_list(df, pais, ed0, ed1, fator)

    all_indicators = []
    for sf in subfactors:
        inds = get_indicator_list(df, pais, ed0, ed1, fator, sf["subfator"])
        all_indicators.extend(inds)

    n_sub_up = sum(1 for item in subfactors if item["status"] == "Melhorou")
    n_sub_down = sum(1 for item in subfactors if item["status"] == "Piorou")
    n_ind_up = sum(1 for item in all_indicators if item["status"] == "Melhorou")
    n_ind_down = sum(1 for item in all_indicators if item["status"] == "Piorou")

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        render_metric_card(n_sub_up, "subfatores melhoraram", "nível intermediário", GREEN)
    with m2:
        render_metric_card(n_sub_down, "subfatores pioraram", "nível intermediário", RED)
    with m3:
        render_metric_card(n_ind_up, "indicadores melhoraram", "nível granular", GREEN)
    with m4:
        render_metric_card(n_ind_down, "indicadores pioraram", "nível granular", RED)

    render_html(
        f"""
        <div class="reading-box">
            <b>Leitura encadeada:</b> cada balão principal representa um <b>subfator</b>.
            Abaixo dele aparecem os <b>indicadores relacionados</b>, conectados visualmente por linhas e marcadores.
        </div>
        """
    )

    indicators_by_subfactor = {
        sf["subfator"]: get_indicator_list(df, pais, ed0, ed1, fator, sf["subfator"])
        for sf in subfactors
    }

    static_html = build_static_factor_html(
        pais=pais,
        ed0=ed0,
        ed1=ed1,
        fator=fator,
        comp=comp,
        subfactors=subfactors,
        indicators_by_subfactor=indicators_by_subfactor,
    )

    st.download_button(
        label="⬇️ Baixar painel estático deste fator (HTML)",
        data=static_html.encode("utf-8"),
        file_name=make_download_filename(fator, pais, ed0, ed1),
        mime="text/html",
        use_container_width=True,
    )

    st.markdown("### Encadeamento completo")

    if not subfactors:
        st.info("Não há subfatores disponíveis para este fator.")
        st.stop()

    for sf in subfactors:
        indicators = get_indicator_list(df, pais, ed0, ed1, fator, sf["subfator"])
        render_subfactor_block(sf, indicators)
