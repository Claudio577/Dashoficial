import os
import re
import unicodedata
from io import BytesIO
from pathlib import Path
from typing import Optional, Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Dashboard Executivo de Suporte",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# ESTILO VISUAL
# ============================================================
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }
        .main-title {
            background: linear-gradient(90deg, #09245f 0%, #083b7c 100%);
            color: white;
            padding: 20px 24px;
            border-radius: 16px;
            margin-bottom: 18px;
            box-shadow: 0 8px 22px rgba(0,0,0,0.12);
        }
        .main-title h1 {
            margin: 0;
            font-size: 30px;
            font-weight: 800;
            letter-spacing: 0.3px;
        }
        .main-title p {
            margin: 4px 0 0 0;
            color: #dbeafe;
            font-size: 15px;
        }
        .kpi-card {
            background: #ffffff;
            border: 1px solid #d9e2f3;
            border-radius: 14px;
            padding: 15px 16px;
            min-height: 126px;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
        }
        .kpi-label {
            color: #0f172a;
            font-size: 13px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: .25px;
            margin-bottom: 7px;
        }
        .kpi-value {
            color: #0f172a;
            font-size: 30px;
            line-height: 34px;
            font-weight: 900;
            margin-bottom: 6px;
        }
        .kpi-meta {
            color: #475569;
            font-size: 12px;
            font-weight: 600;
        }
        .delta-pos { color: #15803d; font-weight: 900; }
        .delta-neg { color: #dc2626; font-weight: 900; }
        .delta-neu { color: #475569; font-weight: 900; }
        .section-title {
            background: #082c63;
            color: #ffffff;
            border-radius: 10px;
            padding: 9px 14px;
            font-size: 15px;
            font-weight: 800;
            text-transform: uppercase;
            margin: 12px 0 10px 0;
        }
        .info-box {
            background: #f8fafc;
            border-left: 5px solid #2563eb;
            border-radius: 12px;
            padding: 14px 16px;
            color: #111827;
            font-size: 14px;
            line-height: 1.45;
            margin-bottom: 10px;
        }
        .alert-box {
            background: #fff7ed;
            border-left: 5px solid #f97316;
            border-radius: 12px;
            padding: 14px 16px;
            color: #111827;
            font-size: 14px;
            line-height: 1.45;
            margin-bottom: 10px;
        }
        .success-box {
            background: #f0fdf4;
            border-left: 5px solid #16a34a;
            border-radius: 12px;
            padding: 14px 16px;
            color: #111827;
            font-size: 14px;
            line-height: 1.45;
            margin-bottom: 10px;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid #d9e2f3;
            border-radius: 12px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
MONTHS_PT = {
    "jan": 1, "janeiro": 1,
    "fev": 2, "fevereiro": 2,
    "mar": 3, "marco": 3, "março": 3,
    "abr": 4, "abril": 4,
    "mai": 5, "maio": 5,
    "jun": 6, "junho": 6,
    "jul": 7, "julho": 7,
    "ago": 8, "agosto": 8,
    "set": 9, "setembro": 9,
    "out": 10, "outubro": 10,
    "nov": 11, "novembro": 11,
    "dez": 12, "dezembro": 12,
}
MONTH_NAMES = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"
}


def strip_accents(text: str) -> str:
    text = str(text)
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def normalize_text(text: str) -> str:
    text = strip_accents(str(text)).lower().strip()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def money_br(value: float) -> str:
    try:
        return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def pct_br(value: float, decimals: int = 1) -> str:
    if pd.isna(value) or np.isinf(value):
        value = 0
    return f"{value:.{decimals}f}%".replace(".", ",")


def number_br(value: float) -> str:
    if pd.isna(value):
        value = 0
    return f"{value:,.0f}".replace(",", ".")


def min_to_human(minutes: float) -> str:
    if pd.isna(minutes) or minutes is None or np.isinf(minutes):
        return "-"
    minutes = max(float(minutes), 0)
    h = int(minutes // 60)
    m = int(minutes % 60)
    s = int(round((minutes - int(minutes)) * 60))
    if h > 0:
        return f"{h}h {m:02d}m"
    return f"{m}m {s:02d}s"


def parse_month_from_filename(filename: str, default_year: int) -> Optional[pd.Timestamp]:
    base = normalize_text(Path(filename).stem)
    for name, month_num in MONTHS_PT.items():
        if re.search(rf"(^|_){normalize_text(name)}(_|$)", base):
            year_match = re.search(r"(20\d{2})", base)
            year = int(year_match.group(1)) if year_match else default_year
            return pd.Timestamp(year=year, month=month_num, day=1)
    return None


def read_excel_any(file_obj, filename: str) -> pd.DataFrame:
    """Lê .xls/.xlsx e tenta ignorar linhas totalmente vazias."""
    try:
        df = pd.read_excel(file_obj, sheet_name=0)
    except Exception:
        # Alguns relatórios antigos têm cabeçalho deslocado. Tenta sem cabeçalho e detecta.
        file_obj.seek(0) if hasattr(file_obj, "seek") else None
        raw = pd.read_excel(file_obj, sheet_name=0, header=None)
        raw = raw.dropna(how="all").dropna(axis=1, how="all")
        header_idx = 0
        best_score = -1
        for i in range(min(20, len(raw))):
            row = raw.iloc[i].astype(str).str.lower().tolist()
            score = sum(any(k in cell for k in ["cliente", "solic", "status", "data", "sla", "assunto", "categoria"]) for cell in row)
            if score > best_score:
                best_score = score
                header_idx = i
        header = raw.iloc[header_idx].tolist()
        df = raw.iloc[header_idx + 1:].copy()
        df.columns = header
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df["arquivo_origem"] = filename
    return df


def load_local_files(data_dir: str = "data") -> list[tuple[str, bytes]]:
    path = Path(data_dir)
    files = []
    if path.exists():
        for p in sorted(path.glob("*.xls*")):
            files.append((p.name, p.read_bytes()))
    return files


def find_column(columns: Iterable[str], keywords: list[str], avoid: list[str] | None = None) -> Optional[str]:
    avoid = avoid or []
    norm_map = {col: normalize_text(col) for col in columns}
    for col, norm in norm_map.items():
        if any(a in norm for a in avoid):
            continue
        if all(k in norm for k in keywords):
            return col
    for col, norm in norm_map.items():
        if any(a in norm for a in avoid):
            continue
        if any(k in norm for k in keywords):
            return col
    return None


def option_index(options: list[str], value: Optional[str]) -> int:
    if value and value in options:
        return options.index(value)
    return 0


def to_datetime_series(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def parse_duration_to_minutes(value) -> float:
    if pd.isna(value):
        return np.nan
    if isinstance(value, pd.Timedelta):
        return value.total_seconds() / 60
    if isinstance(value, (int, float, np.number)):
        # Excel costuma guardar tempo como fração de dia quando valor < 10.
        v = float(value)
        if 0 <= v < 10:
            return v * 24 * 60
        return v

    text = str(value).strip().lower()
    text = strip_accents(text)
    if text in {"", "nan", "none", "nat", "-"}:
        return np.nan

    # HH:MM:SS ou DD HH:MM:SS
    m = re.match(r"(?:(\d+)\s+dias?\s*)?(\d{1,3}):(\d{2})(?::(\d{2}))?$", text)
    if m:
        days = int(m.group(1) or 0)
        h = int(m.group(2) or 0)
        mi = int(m.group(3) or 0)
        se = int(m.group(4) or 0)
        return days * 1440 + h * 60 + mi + se / 60

    total = 0.0
    found = False
    for num, unit in re.findall(r"(\d+(?:[\.,]\d+)?)\s*(dias?|d|horas?|hrs?|hr|h|minutos?|mins?|min|m|segundos?|segs?|seg|s)", text):
        n = float(num.replace(",", "."))
        if unit.startswith("dia") or unit == "d":
            total += n * 1440
        elif unit.startswith("h"):
            total += n * 60
        elif unit.startswith("m"):
            total += n
        elif unit.startswith("s"):
            total += n / 60
        found = True
    if found:
        return total

    # Exemplo: "18m 32s" sem espaço em alguns relatórios
    m2 = re.search(r"(\d+)\s*m", text)
    s2 = re.search(r"(\d+)\s*s", text)
    if m2 or s2:
        return int(m2.group(1)) if m2 else 0 + (int(s2.group(1)) / 60 if s2 else 0)

    return np.nan


def status_closed_mask(series: pd.Series) -> pd.Series:
    s = series.astype(str).map(normalize_text)
    closed_terms = ["fechado", "finalizado", "encerrado", "concluido", "resolvido", "baixado", "atendido"]
    open_terms = ["aberto", "pendente", "andamento", "novo", "aguardando", "pausado"]
    closed = s.apply(lambda x: any(t in x for t in closed_terms))
    opened = s.apply(lambda x: any(t in x for t in open_terms))
    return closed & ~opened


def detect_sla_ok(series: pd.Series) -> pd.Series:
    s = series.astype(str).map(normalize_text)
    ok_terms = ["sim", "ok", "cumprido", "dentro", "no_prazo", "atendido", "verde", "true", "1"]
    bad_terms = ["nao", "fora", "vencido", "estourado", "atrasado", "vermelho", "false", "0"]
    ok = s.apply(lambda x: any(t in x for t in ok_terms))
    bad = s.apply(lambda x: any(t in x for t in bad_terms))
    return ok & ~bad


def detect_reopen(series: pd.Series) -> pd.Series:
    s = series.astype(str).map(normalize_text)
    yes_terms = ["sim", "reab", "recorr", "reincid", "true", "1"]
    no_terms = ["nao", "false", "0"]
    yes = s.apply(lambda x: any(t in x for t in yes_terms))
    no = s.apply(lambda x: any(t in x for t in no_terms))
    return yes & ~no


def build_working_df(df: pd.DataFrame, mapping: dict, default_year: int) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]

    # Data/mês
    open_col = mapping.get("data_abertura")
    close_col = mapping.get("data_fechamento")
    if open_col:
        out["data_abertura_calc"] = to_datetime_series(out[open_col])
    else:
        out["data_abertura_calc"] = pd.NaT

    if close_col:
        out["data_fechamento_calc"] = to_datetime_series(out[close_col])
    else:
        out["data_fechamento_calc"] = pd.NaT

    # Mês pelo arquivo quando não houver data válida.
    file_months = out["arquivo_origem"].apply(lambda x: parse_month_from_filename(str(x), default_year))
    month_from_date = out["data_abertura_calc"].dt.to_period("M").dt.to_timestamp()
    out["mes_ref"] = month_from_date.fillna(file_months)
    out["mes_ref"] = pd.to_datetime(out["mes_ref"], errors="coerce")
    out = out.dropna(subset=["mes_ref"]).copy()
    out["mes_ordem"] = out["mes_ref"].dt.year * 100 + out["mes_ref"].dt.month
    out["mes_nome"] = out["mes_ref"].dt.month.map(MONTH_NAMES) + "/" + out["mes_ref"].dt.year.astype(str)

    # Status
    status_col = mapping.get("status")
    if status_col:
        out["finalizado_calc"] = status_closed_mask(out[status_col])
    else:
        out["finalizado_calc"] = out["data_fechamento_calc"].notna()
    out["aberto_calc"] = ~out["finalizado_calc"]

    # Tempo de atendimento/resolução
    dur_col = mapping.get("tempo_resolucao")
    if dur_col:
        out["tempo_resolucao_min"] = out[dur_col].apply(parse_duration_to_minutes)
    elif open_col and close_col:
        out["tempo_resolucao_min"] = (out["data_fechamento_calc"] - out["data_abertura_calc"]).dt.total_seconds() / 60
    else:
        out["tempo_resolucao_min"] = np.nan

    # SLA
    sla_col = mapping.get("sla")
    sla_limit = mapping.get("sla_limite_min") or 120
    if sla_col:
        out["sla_ok_calc"] = detect_sla_ok(out[sla_col])
    elif out["tempo_resolucao_min"].notna().any():
        out["sla_ok_calc"] = out["tempo_resolucao_min"] <= float(sla_limit)
    else:
        out["sla_ok_calc"] = np.nan
    out["sla_fora_calc"] = out["sla_ok_calc"] == False

    # FCR em até 1 hora
    out["fcr_1h_calc"] = (out["tempo_resolucao_min"] <= 60) & out["finalizado_calc"]

    # Reabertura
    reopen_col = mapping.get("reabertura")
    if reopen_col:
        out["reaberto_calc"] = detect_reopen(out[reopen_col])
    else:
        out["reaberto_calc"] = False

    # Cliente/tema/canal/prioridade
    for target, source in [
        ("cliente_calc", mapping.get("cliente")),
        ("tema_calc", mapping.get("tema")),
        ("canal_calc", mapping.get("canal")),
        ("prioridade_calc", mapping.get("prioridade")),
    ]:
        if source:
            out[target] = out[source].fillna("Não informado").astype(str).str.strip().replace("", "Não informado")
        else:
            out[target] = "Não informado"

    # CSAT
    csat_col = mapping.get("csat")
    if csat_col:
        out["csat_calc"] = pd.to_numeric(out[csat_col].astype(str).str.replace(",", "."), errors="coerce")
    else:
        out["csat_calc"] = np.nan

    return out


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    grp = df.groupby(["mes_ordem", "mes_nome", "mes_ref"], as_index=False).agg(
        total=("arquivo_origem", "count"),
        abertos=("aberto_calc", "sum"),
        finalizados=("finalizado_calc", "sum"),
        sla_ok=("sla_ok_calc", lambda x: int(pd.Series(x).fillna(False).sum())),
        sla_fora=("sla_fora_calc", lambda x: int(pd.Series(x).fillna(False).sum())),
        fcr_1h=("fcr_1h_calc", "sum"),
        reabertos=("reaberto_calc", "sum"),
        tma_min=("tempo_resolucao_min", "mean"),
        csat=("csat_calc", "mean"),
    )
    grp["perc_sla"] = np.where(grp["total"] > 0, grp["sla_ok"] / grp["total"] * 100, 0)
    grp["perc_fcr"] = np.where(grp["finalizados"] > 0, grp["fcr_1h"] / grp["finalizados"] * 100, 0)
    grp["perc_reabertura"] = np.where(grp["total"] > 0, grp["reabertos"] / grp["total"] * 100, 0)
    return grp.sort_values("mes_ordem")


def calc_delta(curr, prev, inverse: bool = False, suffix: str = "") -> tuple[str, str]:
    if prev is None or pd.isna(prev) or float(prev) == 0:
        return "sem mês anterior", "delta-neu"
    diff = float(curr) - float(prev)
    pct = diff / abs(float(prev)) * 100
    arrow = "▲" if diff > 0 else "▼" if diff < 0 else "—"
    good = diff >= 0
    if inverse:
        good = diff <= 0
    cls = "delta-pos" if good and diff != 0 else "delta-neg" if diff != 0 else "delta-neu"
    text = f"{arrow} {pct:+.1f}% vs mês anterior".replace(".", ",")
    if suffix:
        text += f" {suffix}"
    return text, cls


def kpi_card(label: str, value: str, meta: str, delta_text: str = "", delta_cls: str = "delta-neu"):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-meta">{meta}</div>
            <div class="{delta_cls}">{delta_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def safe_bar(df, x, y, title, text=None):
    fig = px.bar(df, x=x, y=y, text=text, title=title)
    fig.update_traces(textposition="outside")
    fig.update_layout(
        height=360,
        margin=dict(l=20, r=20, t=60, b=20),
        title_font=dict(size=16, color="#0f172a"),
        font=dict(color="#0f172a", size=12),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        yaxis_title=None,
        xaxis_title=None,
    )
    return fig


def safe_line(df, x, y, title):
    fig = px.line(df, x=x, y=y, markers=True, title=title)
    fig.update_layout(
        height=360,
        margin=dict(l=20, r=20, t=60, b=20),
        title_font=dict(size=16, color="#0f172a"),
        font=dict(color="#0f172a", size=12),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        yaxis_title=None,
        xaxis_title=None,
    )
    return fig


# ============================================================
# CABEÇALHO
# ============================================================
st.markdown(
    """
    <div class="main-title">
        <h1>SUPORTE — VISÃO EXECUTIVA PARA TOMADA DE DECISÃO</h1>
        <p>Indicadores anuais, comparativo mensal, SLA, FCR, backlog, clientes críticos e pontos de atenção.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# CARREGAMENTO DOS DADOS
# ============================================================
st.sidebar.header("📁 Base de dados")
st.sidebar.caption("Use arquivos .xls ou .xlsx, um por mês ou uma base única anual.")

uploaded = st.sidebar.file_uploader(
    "Enviar relatórios mensais",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
)

local_files = load_local_files("data")
all_file_bytes: list[tuple[str, bytes]] = []
all_file_bytes.extend(local_files)
if uploaded:
    all_file_bytes.extend([(f.name, f.getvalue()) for f in uploaded])

if not all_file_bytes:
    st.info(
        "Coloque os arquivos de chamados dentro da pasta `data/` no GitHub ou envie os arquivos pela barra lateral."
    )
    st.stop()

loaded_frames = []
load_errors = []
for filename, content in all_file_bytes:
    try:
        frame = read_excel_any(BytesIO(content), filename)
        loaded_frames.append(frame)
    except Exception as e:
        load_errors.append((filename, str(e)))

if load_errors:
    st.warning("Alguns arquivos não foram carregados:")
    for name, err in load_errors:
        st.write(f"- **{name}**: {err}")

if not loaded_frames:
    st.error("Nenhum arquivo foi carregado com sucesso.")
    st.stop()

raw_df = pd.concat(loaded_frames, ignore_index=True, sort=False)
raw_df.columns = [str(c).strip() for c in raw_df.columns]

# ============================================================
# MAPEAMENTO DAS COLUNAS
# ============================================================
all_columns = list(raw_df.columns)
select_options = ["Não usar"] + all_columns

auto_open = find_column(all_columns, ["abertura"]) or find_column(all_columns, ["data"], avoid=["fech", "fim", "encerr", "conclus"])
auto_close = find_column(all_columns, ["fech"]) or find_column(all_columns, ["encerr"]) or find_column(all_columns, ["conclus"])
auto_status = find_column(all_columns, ["status"]) or find_column(all_columns, ["situacao"])
auto_cliente = find_column(all_columns, ["cliente"]) or find_column(all_columns, ["empresa"]) or find_column(all_columns, ["solicitante"])
auto_tema = find_column(all_columns, ["categoria"]) or find_column(all_columns, ["assunto"]) or find_column(all_columns, ["motivo"]) or find_column(all_columns, ["tipo"])
auto_sla = find_column(all_columns, ["sla"]) or find_column(all_columns, ["prazo"])
auto_dur = find_column(all_columns, ["tempo", "solucao"]) or find_column(all_columns, ["tempo", "atendimento"]) or find_column(all_columns, ["duracao"])
auto_prior = find_column(all_columns, ["prioridade"]) or find_column(all_columns, ["criticidade"])
auto_reopen = find_column(all_columns, ["reab"]) or find_column(all_columns, ["reincid"])
auto_csat = find_column(all_columns, ["csat"]) or find_column(all_columns, ["satisfacao"]) or find_column(all_columns, ["nota"])
auto_channel = find_column(all_columns, ["canal"]) or find_column(all_columns, ["origem"])

with st.sidebar.expander("⚙️ Mapeamento das colunas", expanded=True):
    data_abertura = st.selectbox("Data de abertura", select_options, index=option_index(select_options, auto_open))
    data_fechamento = st.selectbox("Data de fechamento/conclusão", select_options, index=option_index(select_options, auto_close))
    status_col = st.selectbox("Status", select_options, index=option_index(select_options, auto_status))
    cliente_col = st.selectbox("Cliente / empresa", select_options, index=option_index(select_options, auto_cliente))
    tema_col = st.selectbox("Assunto / categoria / motivo", select_options, index=option_index(select_options, auto_tema))
    sla_col = st.selectbox("SLA", select_options, index=option_index(select_options, auto_sla))
    tempo_col = st.selectbox("Tempo de resolução/atendimento", select_options, index=option_index(select_options, auto_dur))
    prioridade_col = st.selectbox("Prioridade", select_options, index=option_index(select_options, auto_prior))
    canal_col = st.selectbox("Canal", select_options, index=option_index(select_options, auto_channel))
    reabertura_col = st.selectbox("Reabertura/reincidência", select_options, index=option_index(select_options, auto_reopen))
    csat_col = st.selectbox("CSAT/satisfação", select_options, index=option_index(select_options, auto_csat))

with st.sidebar.expander("🎯 Metas", expanded=False):
    default_year = st.number_input("Ano padrão quando o mês vier só no nome do arquivo", min_value=2020, max_value=2035, value=pd.Timestamp.today().year)
    meta_sla = st.number_input("Meta de SLA cumprido (%)", min_value=0, max_value=100, value=90)
    meta_fcr = st.number_input("Meta de FCR 1h (%)", min_value=0, max_value=100, value=70)
    meta_tma_min = st.number_input("Meta de TMA em minutos", min_value=1, max_value=10000, value=20)
    sla_limite_min = st.number_input("Limite usado para calcular SLA quando não existir coluna SLA (min)", min_value=1, max_value=10000, value=120)

mapping = {
    "data_abertura": None if data_abertura == "Não usar" else data_abertura,
    "data_fechamento": None if data_fechamento == "Não usar" else data_fechamento,
    "status": None if status_col == "Não usar" else status_col,
    "cliente": None if cliente_col == "Não usar" else cliente_col,
    "tema": None if tema_col == "Não usar" else tema_col,
    "sla": None if sla_col == "Não usar" else sla_col,
    "tempo_resolucao": None if tempo_col == "Não usar" else tempo_col,
    "prioridade": None if prioridade_col == "Não usar" else prioridade_col,
    "canal": None if canal_col == "Não usar" else canal_col,
    "reabertura": None if reabertura_col == "Não usar" else reabertura_col,
    "csat": None if csat_col == "Não usar" else csat_col,
    "sla_limite_min": sla_limite_min,
}

df = build_working_df(raw_df, mapping, int(default_year))

if df.empty:
    st.error("Não foi possível identificar o mês de referência. Use data de abertura ou arquivos com mês no nome, como abril.xls, maio.xls.")
    st.stop()

summary = monthly_summary(df)

# Filtro de meses
month_options = summary["mes_nome"].tolist()
selected_months = st.sidebar.multiselect("Filtrar meses", month_options, default=month_options)
if selected_months:
    df = df[df["mes_nome"].isin(selected_months)]
    summary = monthly_summary(df)

if summary.empty:
    st.warning("Nenhum dado disponível para os filtros selecionados.")
    st.stop()

current = summary.iloc[-1]
previous = summary.iloc[-2] if len(summary) >= 2 else None

st.caption(
    f"Arquivos carregados: **{len(all_file_bytes)}** | Linhas analisadas: **{number_br(len(df))}** | Período: **{summary.iloc[0]['mes_nome']} a {summary.iloc[-1]['mes_nome']}**"
)

# ============================================================
# KPIS
# ============================================================
st.markdown('<div class="section-title">Indicadores principais do mês atual</div>', unsafe_allow_html=True)

c1, c2, c3, c4, c5, c6 = st.columns(6)

prev_total = previous["total"] if previous is not None else None
prev_sla = previous["perc_sla"] if previous is not None else None
prev_tma = previous["tma_min"] if previous is not None else None
prev_fcr = previous["perc_fcr"] if previous is not None else None
prev_backlog = previous["abertos"] if previous is not None else None
prev_reopen = previous["perc_reabertura"] if previous is not None else None

with c1:
    d, cls = calc_delta(current["total"], prev_total, inverse=True)
    kpi_card("Chamados", number_br(current["total"]), f"Mês atual: {current['mes_nome']}", d, cls)
with c2:
    d, cls = calc_delta(current["perc_sla"], prev_sla, inverse=False)
    kpi_card("SLA cumprido", pct_br(current["perc_sla"], 1), f"Meta: {meta_sla}%", d, cls)
with c3:
    d, cls = calc_delta(current["tma_min"], prev_tma, inverse=True)
    kpi_card("TMA", min_to_human(current["tma_min"]), f"Meta: até {meta_tma_min} min", d, cls)
with c4:
    d, cls = calc_delta(current["perc_fcr"], prev_fcr, inverse=False)
    kpi_card("FCR 1h", pct_br(current["perc_fcr"], 1), f"Meta: {meta_fcr}%", d, cls)
with c5:
    d, cls = calc_delta(current["abertos"], prev_backlog, inverse=True)
    kpi_card("Backlog", number_br(current["abertos"]), "Chamados ainda abertos", d, cls)
with c6:
    d, cls = calc_delta(current["perc_reabertura"], prev_reopen, inverse=True)
    kpi_card("Reabertura", pct_br(current["perc_reabertura"], 1), "Reincidência/reabertos", d, cls)

# ============================================================
# GRÁFICOS PRINCIPAIS
# ============================================================
st.markdown('<div class="section-title">Evolução mensal e comparativos</div>', unsafe_allow_html=True)

summary_chart = summary.copy()
summary_chart["SLA %"] = summary_chart["perc_sla"].round(1)
summary_chart["FCR 1h %"] = summary_chart["perc_fcr"].round(1)
summary_chart["TMA (min)"] = summary_chart["tma_min"].round(1)

col_a, col_b = st.columns(2)
with col_a:
    fig = safe_bar(summary_chart, "mes_nome", "total", "Volume de chamados por mês", text="total")
    st.plotly_chart(fig, use_container_width=True)
with col_b:
    sla_fig = safe_line(summary_chart, "mes_nome", "SLA %", "SLA cumprido — evolução mensal (%)")
    sla_fig.add_hline(y=meta_sla, line_dash="dash", annotation_text=f"Meta {meta_sla}%")
    st.plotly_chart(sla_fig, use_container_width=True)

col_c, col_d = st.columns(2)
with col_c:
    fcr_fig = safe_line(summary_chart, "mes_nome", "FCR 1h %", "FCR em até 1 hora — evolução mensal (%)")
    fcr_fig.add_hline(y=meta_fcr, line_dash="dash", annotation_text=f"Meta {meta_fcr}%")
    st.plotly_chart(fcr_fig, use_container_width=True)
with col_d:
    comp = summary_chart[["mes_nome", "abertos", "finalizados", "sla_fora"]].melt(
        id_vars="mes_nome", var_name="Indicador", value_name="Quantidade"
    )
    fig = px.bar(comp, x="mes_nome", y="Quantidade", color="Indicador", barmode="group", title="Abertos x finalizados x fora do SLA")
    fig.update_layout(height=360, margin=dict(l=20, r=20, t=60, b=20), font=dict(color="#0f172a", size=12))
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# BACKLOG E TOPS
# ============================================================
st.markdown('<div class="section-title">Backlog, clientes e temas críticos</div>', unsafe_allow_html=True)

col_e, col_f, col_g = st.columns([1, 1, 1])
with col_e:
    curr_df = df[df["mes_nome"] == current["mes_nome"]]
    backlog_labels = pd.DataFrame({
        "Status": ["Dentro do SLA", "Fora do SLA"],
        "Quantidade": [int((curr_df["aberto_calc"] & curr_df["sla_ok_calc"].fillna(False)).sum()), int((curr_df["aberto_calc"] & curr_df["sla_fora_calc"].fillna(False)).sum())]
    })
    if backlog_labels["Quantidade"].sum() == 0:
        backlog_labels = pd.DataFrame({"Status": ["Sem backlog classificado"], "Quantidade": [1]})
    fig = px.pie(backlog_labels, names="Status", values="Quantidade", hole=0.55, title="Backlog — dentro x fora do SLA")
    fig.update_layout(height=360, margin=dict(l=20, r=20, t=60, b=20), font=dict(color="#0f172a", size=12))
    st.plotly_chart(fig, use_container_width=True)

with col_f:
    top_clientes = curr_df["cliente_calc"].value_counts().head(8).reset_index()
    top_clientes.columns = ["Cliente", "Chamados"]
    fig = px.bar(top_clientes.sort_values("Chamados"), x="Chamados", y="Cliente", orientation="h", title="Top clientes do mês")
    fig.update_layout(height=360, margin=dict(l=20, r=20, t=60, b=20), font=dict(color="#0f172a", size=12), yaxis_title=None, xaxis_title=None)
    st.plotly_chart(fig, use_container_width=True)

with col_g:
    top_temas = curr_df["tema_calc"].value_counts().head(8).reset_index()
    top_temas.columns = ["Tema", "Chamados"]
    fig = px.bar(top_temas.sort_values("Chamados"), x="Chamados", y="Tema", orientation="h", title="Top temas/motivos do mês")
    fig.update_layout(height=360, margin=dict(l=20, r=20, t=60, b=20), font=dict(color="#0f172a", size=12), yaxis_title=None, xaxis_title=None)
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# ANÁLISE EXECUTIVA
# ============================================================
st.markdown('<div class="section-title">Leitura executiva automática</div>', unsafe_allow_html=True)

alerts = []
success = []
recs = []

if current["perc_sla"] < meta_sla:
    alerts.append(f"SLA abaixo da meta: {pct_br(current['perc_sla'])}, meta de {meta_sla}%.")
    recs.append("Priorizar chamados fora do SLA e revisar gargalos por cliente, tema e equipe.")
else:
    success.append(f"SLA dentro/acima da meta: {pct_br(current['perc_sla'])}.")

if current["perc_fcr"] < meta_fcr:
    alerts.append(f"FCR 1h abaixo da meta: {pct_br(current['perc_fcr'])}, meta de {meta_fcr}%.")
    recs.append("Criar base de conhecimento, scripts de atendimento e automações para resolução no primeiro contato.")
else:
    success.append(f"FCR 1h dentro/acima da meta: {pct_br(current['perc_fcr'])}.")

if pd.notna(current["tma_min"]) and current["tma_min"] > meta_tma_min:
    alerts.append(f"TMA acima da meta: {min_to_human(current['tma_min'])}, meta de até {meta_tma_min} min.")
    recs.append("Analisar filas, motivos mais recorrentes e volume por dia para reduzir tempo médio de atendimento.")
else:
    success.append(f"TMA dentro da meta ou sem dado suficiente para alerta: {min_to_human(current['tma_min'])}.")

if current["abertos"] > 0:
    alerts.append(f"Backlog atual: {number_br(current['abertos'])} chamado(s) aberto(s).")
    recs.append("Acompanhar backlog diariamente e separar chamados críticos dos chamados de baixa prioridade.")

if previous is not None:
    if current["total"] > previous["total"]:
        alerts.append(f"Volume aumentou em relação ao mês anterior: {number_br(previous['total'])} → {number_br(current['total'])}.")
    elif current["total"] < previous["total"]:
        success.append(f"Volume reduziu em relação ao mês anterior: {number_br(previous['total'])} → {number_br(current['total'])}.")

col_h, col_i, col_j = st.columns(3)
with col_h:
    st.markdown("<div class='alert-box'><b>Riscos / pontos de atenção</b><br>" + "<br>".join([f"• {a}" for a in alerts[:6]]) + "</div>", unsafe_allow_html=True)
with col_i:
    st.markdown("<div class='success-box'><b>Evoluções positivas</b><br>" + "<br>".join([f"• {s}" for s in success[:6]]) + "</div>", unsafe_allow_html=True)
with col_j:
    st.markdown("<div class='info-box'><b>Recomendações</b><br>" + "<br>".join([f"• {r}" for r in list(dict.fromkeys(recs))[:6]]) + "</div>", unsafe_allow_html=True)

# ============================================================
# TABELAS
# ============================================================
st.markdown('<div class="section-title">Tabelas de apoio</div>', unsafe_allow_html=True)

summary_view = summary.copy()
summary_view = summary_view[[
    "mes_nome", "total", "abertos", "finalizados", "sla_ok", "sla_fora", "perc_sla", "fcr_1h", "perc_fcr", "tma_min", "reabertos", "perc_reabertura", "csat"
]]
summary_view.columns = [
    "Mês", "Total", "Abertos", "Finalizados", "SLA cumprido", "Fora do SLA", "% SLA", "FCR 1h", "% FCR 1h", "TMA (min)", "Reabertos", "% Reabertura", "CSAT"
]

st.dataframe(
    summary_view.style.format({
        "% SLA": "{:.1f}%",
        "% FCR 1h": "{:.1f}%",
        "TMA (min)": "{:.1f}",
        "% Reabertura": "{:.1f}%",
        "CSAT": "{:.2f}",
    }),
    use_container_width=True,
    hide_index=True,
)

with st.expander("Ver base consolidada tratada"):
    st.dataframe(df, use_container_width=True, hide_index=True)

csv = df.to_csv(index=False, sep=";", encoding="utf-8-sig")
st.download_button(
    "⬇️ Baixar base consolidada tratada em CSV",
    data=csv.encode("utf-8-sig"),
    file_name="base_consolidada_suporte.csv",
    mime="text/csv",
)
