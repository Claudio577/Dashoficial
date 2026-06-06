import os
import re
import unicodedata
from io import BytesIO
from pathlib import Path
from typing import Iterable, Optional

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# ==========================================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================================
st.set_page_config(
    page_title="Dashboard Executivo de Suporte",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==========================================================
# ESTILO VISUAL
# ==========================================================
st.markdown(
    """
    <style>
        .block-container { padding-top: 1.1rem; padding-bottom: 2rem; }
        .main-title {
            background: linear-gradient(90deg, #09245f 0%, #073b7c 100%);
            color: white;
            padding: 20px 24px;
            border-radius: 16px;
            margin-bottom: 16px;
            box-shadow: 0 8px 22px rgba(0,0,0,0.12);
        }
        .main-title h1 { margin: 0; font-size: 30px; font-weight: 900; letter-spacing: .3px; }
        .main-title p { margin: 5px 0 0 0; color: #dbeafe; font-size: 15px; }
        .section-title {
            background: #082c63;
            color: white;
            border-radius: 10px;
            padding: 9px 14px;
            font-size: 15px;
            font-weight: 900;
            text-transform: uppercase;
            margin: 14px 0 10px 0;
        }
        .kpi-card {
            background: #ffffff;
            border: 1px solid #d9e2f3;
            border-radius: 15px;
            padding: 14px 15px;
            min-height: 126px;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
        }
        .kpi-label {
            color: #0f172a;
            font-size: 12.5px;
            font-weight: 900;
            text-transform: uppercase;
            letter-spacing: .25px;
            margin-bottom: 8px;
        }
        .kpi-value {
            color: #0f172a;
            font-size: 29px;
            line-height: 34px;
            font-weight: 950;
            margin-bottom: 5px;
        }
        .kpi-meta { color: #475569; font-size: 12px; font-weight: 700; }
        .delta-pos { color: #15803d; font-weight: 900; font-size: 12px; margin-top: 6px; }
        .delta-neg { color: #dc2626; font-weight: 900; font-size: 12px; margin-top: 6px; }
        .delta-neu { color: #475569; font-weight: 900; font-size: 12px; margin-top: 6px; }
        .info-box, .alert-box, .success-box {
            border-radius: 12px;
            padding: 14px 16px;
            color: #111827;
            font-size: 14px;
            line-height: 1.45;
            min-height: 172px;
            border: 1px solid #e5e7eb;
        }
        .info-box { background: #eff6ff; border-left: 5px solid #2563eb; }
        .alert-box { background: #fff7ed; border-left: 5px solid #f97316; }
        .success-box { background: #f0fdf4; border-left: 5px solid #16a34a; }
        div[data-testid="stDataFrame"] { border: 1px solid #d9e2f3; border-radius: 12px; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================================
# CONSTANTES E FUNÇÕES AUXILIARES
# ==========================================================
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
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
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


def number_br(value) -> str:
    try:
        if pd.isna(value):
            value = 0
        return f"{float(value):,.0f}".replace(",", ".")
    except Exception:
        return "0"


def pct_br(value, decimals: int = 1) -> str:
    try:
        if pd.isna(value) or np.isinf(value):
            value = 0
        return f"{float(value):.{decimals}f}%".replace(".", ",")
    except Exception:
        return "0,0%"


def min_to_human(minutes) -> str:
    try:
        if pd.isna(minutes) or np.isinf(minutes):
            return "-"
        minutes = max(float(minutes), 0)
        h = int(minutes // 60)
        m = int(minutes % 60)
        s = int(round((minutes - int(minutes)) * 60))
        if h > 0:
            return f"{h}h {m:02d}m"
        return f"{m}m {s:02d}s"
    except Exception:
        return "-"


def make_unique_columns(columns: Iterable) -> list[str]:
    seen = {}
    result = []
    for i, col in enumerate(columns):
        name = str(col).strip()
        if name == "" or name.lower() in {"nan", "none", "unnamed: 0"}:
            name = f"coluna_{i + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        result.append(name)
    return result


def excel_engine(filename: str) -> Optional[str]:
    name = filename.lower()
    if name.endswith(".xlsx"):
        return "openpyxl"
    if name.endswith(".xls"):
        return "xlrd"
    return None


def detect_header_row(raw: pd.DataFrame) -> int:
    keywords = [
        "cliente", "empresa", "solicitante", "chamado", "ticket", "status",
        "situacao", "abertura", "fechamento", "conclusao", "sla", "prazo",
        "assunto", "categoria", "motivo", "prioridade", "tempo", "canal",
    ]
    best_idx = 0
    best_score = -1
    max_rows = min(30, len(raw))
    for i in range(max_rows):
        values = [normalize_text(v) for v in raw.iloc[i].tolist()]
        non_empty = sum(v not in {"", "nan", "none", "nat"} for v in values)
        keyword_hits = sum(any(k in v for k in keywords) for v in values)
        score = keyword_hits * 10 + non_empty * 0.1
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx


def read_excel_with_header_detection(content: bytes, filename: str, sheet_name) -> pd.DataFrame:
    engine = excel_engine(filename)
    raw = pd.read_excel(
        BytesIO(content),
        sheet_name=sheet_name,
        header=None,
        engine=engine,
    )
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        return pd.DataFrame()

    header_idx = detect_header_row(raw)
    header = make_unique_columns(raw.iloc[header_idx].tolist())
    df = raw.iloc[header_idx + 1:].copy()
    df.columns = header
    df = df.dropna(how="all").dropna(axis=1, how="all")
    return df


def read_excel_file(content: bytes, filename: str) -> pd.DataFrame:
    engine = excel_engine(filename)
    try:
        excel = pd.ExcelFile(BytesIO(content), engine=engine)
        sheets = excel.sheet_names
    except Exception:
        sheets = [0]

    frames = []
    for sheet in sheets:
        try:
            df_sheet = read_excel_with_header_detection(content, filename, sheet)
            if df_sheet.empty:
                continue
            df_sheet["arquivo_origem"] = filename
            df_sheet["aba_origem"] = str(sheet)
            frames.append(df_sheet)
        except Exception:
            continue

    if not frames:
        raise ValueError("Não foi possível ler nenhuma aba desse arquivo.")

    return pd.concat(frames, ignore_index=True, sort=False)


def find_column(columns: Iterable[str], keywords: list[str], avoid: Optional[list[str]] = None) -> Optional[str]:
    avoid = avoid or []
    norm_map = {col: normalize_text(col) for col in columns}

    # Primeiro tenta encontrar todas as palavras-chave na coluna.
    for col, norm in norm_map.items():
        if any(a in norm for a in avoid):
            continue
        if all(k in norm for k in keywords):
            return col

    # Depois aceita encontrar pelo menos uma palavra-chave.
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


def parse_month_from_filename(filename: str, default_year: int) -> Optional[pd.Timestamp]:
    base = normalize_text(Path(filename).stem)
    for name, month_num in MONTHS_PT.items():
        name_norm = normalize_text(name)
        if re.search(rf"(^|_){name_norm}(_|$)", base):
            year_match = re.search(r"(20\d{2})", base)
            year = int(year_match.group(1)) if year_match else int(default_year)
            return pd.Timestamp(year=year, month=month_num, day=1)
    return None


def to_datetime_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def parse_duration_to_minutes(value) -> float:
    if pd.isna(value):
        return np.nan

    if isinstance(value, pd.Timedelta):
        return value.total_seconds() / 60

    if isinstance(value, (int, float, np.number)):
        v = float(value)
        # Excel pode armazenar tempo como fração de dia. Ex.: 0,5 = 12 horas.
        if 0 <= v < 10:
            return v * 24 * 60
        return v

    text = strip_accents(str(value)).lower().strip()
    if text in {"", "nan", "none", "nat", "-"}:
        return np.nan

    # Formatos: HH:MM:SS, H:MM, 2 dias 03:10:00
    m = re.match(r"(?:(\d+)\s+dias?\s*)?(\d{1,3}):(\d{2})(?::(\d{2}))?$", text)
    if m:
        days = int(m.group(1) or 0)
        hours = int(m.group(2) or 0)
        minutes = int(m.group(3) or 0)
        seconds = int(m.group(4) or 0)
        return days * 1440 + hours * 60 + minutes + seconds / 60

    # Formatos: 1d 2h 30m, 18m 32s, 4 horas, 30 min
    total = 0.0
    found = False
    pattern = r"(\d+(?:[\.,]\d+)?)\s*(dias?|dia|d|horas?|hrs?|hr|h|minutos?|mins?|min|m|segundos?|segs?|seg|s)"
    for num, unit in re.findall(pattern, text):
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


def build_working_df(raw_df: pd.DataFrame, mapping: dict, default_year: int) -> pd.DataFrame:
    df = raw_df.copy()
    df.columns = make_unique_columns(df.columns)

    open_col = mapping.get("data_abertura")
    close_col = mapping.get("data_fechamento")
    status_col = mapping.get("status")
    tempo_col = mapping.get("tempo_resolucao")
    sla_col = mapping.get("sla")
    sla_limit = float(mapping.get("sla_limite_min") or 120)

    if open_col:
        df["data_abertura_calc"] = to_datetime_series(df[open_col])
    else:
        df["data_abertura_calc"] = pd.NaT

    if close_col:
        df["data_fechamento_calc"] = to_datetime_series(df[close_col])
    else:
        df["data_fechamento_calc"] = pd.NaT

    # Primeiro tenta usar a data de abertura; se não tiver, usa o mês no nome do arquivo.
    file_months = df["arquivo_origem"].apply(lambda x: parse_month_from_filename(str(x), default_year))
    month_from_date = df["data_abertura_calc"].dt.to_period("M").dt.to_timestamp()
    df["mes_ref"] = month_from_date.fillna(file_months)
    df["mes_ref"] = pd.to_datetime(df["mes_ref"], errors="coerce")
    df = df.dropna(subset=["mes_ref"]).copy()

    df["mes_ordem"] = df["mes_ref"].dt.year * 100 + df["mes_ref"].dt.month
    df["mes_nome"] = df["mes_ref"].dt.month.map(MONTH_NAMES) + "/" + df["mes_ref"].dt.year.astype(str)

    if status_col:
        df["finalizado_calc"] = status_closed_mask(df[status_col])
    else:
        df["finalizado_calc"] = df["data_fechamento_calc"].notna()
    df["aberto_calc"] = ~df["finalizado_calc"]

    if tempo_col:
        df["tempo_resolucao_min"] = df[tempo_col].apply(parse_duration_to_minutes)
    elif open_col and close_col:
        df["tempo_resolucao_min"] = (
            df["data_fechamento_calc"] - df["data_abertura_calc"]
        ).dt.total_seconds() / 60
    else:
        df["tempo_resolucao_min"] = np.nan

    if sla_col:
        df["sla_ok_calc"] = detect_sla_ok(df[sla_col])
    elif df["tempo_resolucao_min"].notna().any():
        df["sla_ok_calc"] = df["tempo_resolucao_min"] <= sla_limit
    else:
        df["sla_ok_calc"] = np.nan
    df["sla_fora_calc"] = df["sla_ok_calc"] == False

    df["fcr_1h_calc"] = (df["tempo_resolucao_min"] <= 60) & df["finalizado_calc"]

    reopen_col = mapping.get("reabertura")
    if reopen_col:
        df["reaberto_calc"] = detect_reopen(df[reopen_col])
    else:
        df["reaberto_calc"] = False

    for target, source in [
        ("cliente_calc", mapping.get("cliente")),
        ("tema_calc", mapping.get("tema")),
        ("canal_calc", mapping.get("canal")),
        ("prioridade_calc", mapping.get("prioridade")),
    ]:
        if source:
            df[target] = df[source].fillna("Não informado").astype(str).str.strip()
            df[target] = df[target].replace("", "Não informado")
        else:
            df[target] = "Não informado"

    csat_col = mapping.get("csat")
    if csat_col:
        df["csat_calc"] = pd.to_numeric(
            df[csat_col].astype(str).str.replace(",", ".", regex=False),
            errors="coerce",
        )
    else:
        df["csat_calc"] = np.nan

    return df


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    summary = df.groupby(["mes_ordem", "mes_nome", "mes_ref"], as_index=False).agg(
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
    summary["perc_sla"] = np.where(summary["total"] > 0, summary["sla_ok"] / summary["total"] * 100, 0)
    summary["perc_fcr"] = np.where(summary["finalizados"] > 0, summary["fcr_1h"] / summary["finalizados"] * 100, 0)
    summary["perc_reabertura"] = np.where(summary["total"] > 0, summary["reabertos"] / summary["total"] * 100, 0)
    return summary.sort_values("mes_ordem")


def calc_delta(curr, prev, inverse: bool = False) -> tuple[str, str]:
    if prev is None or pd.isna(prev) or float(prev) == 0:
        return "sem mês anterior", "delta-neu"
    diff = float(curr) - float(prev)
    pct = diff / abs(float(prev)) * 100
    arrow = "▲" if diff > 0 else "▼" if diff < 0 else "—"
    good = diff >= 0
    if inverse:
        good = diff <= 0
    css_class = "delta-pos" if good and diff != 0 else "delta-neg" if diff != 0 else "delta-neu"
    text = f"{arrow} {pct:+.1f}% vs mês anterior".replace(".", ",")
    return text, css_class


def kpi_card(label: str, value: str, meta: str, delta_text: str, delta_cls: str):
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


def base_fig_layout(fig, height: int = 360):
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=60, b=20),
        title_font=dict(size=16, color="#0f172a"),
        font=dict(color="#0f172a", size=12),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        xaxis_title=None,
        yaxis_title=None,
    )
    return fig


# ==========================================================
# CABEÇALHO
# ==========================================================
st.markdown(
    """
    <div class="main-title">
        <h1>SUPORTE — VISÃO EXECUTIVA PARA TOMADA DE DECISÃO</h1>
        <p>Indicadores mensais e anuais com SLA, FCR 1h, TMA, backlog, clientes, temas críticos e recomendações.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ==========================================================
# UPLOAD DOS ARQUIVOS
# ==========================================================
st.sidebar.header("📁 Base de dados")
st.sidebar.caption("Envie arquivos .xls ou .xlsx. Pode enviar abril, maio e depois o ano todo.")

uploaded_files = st.sidebar.file_uploader(
    "Enviar relatórios mensais",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info(
        "Envie os arquivos Excel pela barra lateral. Não precisa colocar os Excel no GitHub. "
        "O GitHub precisa ter apenas app.py, requirements.txt, README.md e .gitignore."
    )
    st.stop()

loaded_frames = []
load_errors = []
for file in uploaded_files:
    try:
        content = file.getvalue()
        frame = read_excel_file(content, file.name)
        loaded_frames.append(frame)
    except Exception as exc:
        load_errors.append((file.name, str(exc)))

if load_errors:
    st.warning("Alguns arquivos não foram carregados:")
    for name, err in load_errors:
        st.write(f"- **{name}**: {err}")

if not loaded_frames:
    st.error("Nenhum arquivo foi carregado com sucesso. Confira se o requirements.txt tem xlrd e openpyxl.")
    st.stop()

raw_df = pd.concat(loaded_frames, ignore_index=True, sort=False)
raw_df.columns = make_unique_columns(raw_df.columns)


# ==========================================================
# MAPEAMENTO DAS COLUNAS
# ==========================================================
all_columns = list(raw_df.columns)
select_options = ["Não usar"] + all_columns

auto_open = (
    find_column(all_columns, ["abertura"])
    or find_column(all_columns, ["criacao"])
    or find_column(all_columns, ["data"], avoid=["fech", "fim", "encerr", "conclus", "solucao"])
)
auto_close = (
    find_column(all_columns, ["fech"])
    or find_column(all_columns, ["encerr"])
    or find_column(all_columns, ["conclus"])
    or find_column(all_columns, ["solucao"])
)
auto_status = find_column(all_columns, ["status"]) or find_column(all_columns, ["situacao"])
auto_cliente = (
    find_column(all_columns, ["cliente"])
    or find_column(all_columns, ["empresa"])
    or find_column(all_columns, ["solicitante"])
    or find_column(all_columns, ["nome"])
)
auto_tema = (
    find_column(all_columns, ["categoria"])
    or find_column(all_columns, ["assunto"])
    or find_column(all_columns, ["motivo"])
    or find_column(all_columns, ["tipo"])
    or find_column(all_columns, ["descricao"])
)
auto_sla = find_column(all_columns, ["sla"]) or find_column(all_columns, ["prazo"])
auto_tempo = (
    find_column(all_columns, ["tempo", "solucao"])
    or find_column(all_columns, ["tempo", "atendimento"])
    or find_column(all_columns, ["tempo", "resolucao"])
    or find_column(all_columns, ["duracao"])
)
auto_prioridade = find_column(all_columns, ["prioridade"]) or find_column(all_columns, ["criticidade"])
auto_canal = find_column(all_columns, ["canal"]) or find_column(all_columns, ["origem"])
auto_reabertura = find_column(all_columns, ["reab"]) or find_column(all_columns, ["reincid"])
auto_csat = find_column(all_columns, ["csat"]) or find_column(all_columns, ["satisfacao"]) or find_column(all_columns, ["nota"])

with st.sidebar.expander("⚙️ Mapeamento das colunas", expanded=True):
    data_abertura = st.selectbox("Data de abertura", select_options, index=option_index(select_options, auto_open))
    data_fechamento = st.selectbox("Data de fechamento/conclusão", select_options, index=option_index(select_options, auto_close))
    status_col = st.selectbox("Status", select_options, index=option_index(select_options, auto_status))
    cliente_col = st.selectbox("Cliente / empresa", select_options, index=option_index(select_options, auto_cliente))
    tema_col = st.selectbox("Assunto / categoria / motivo", select_options, index=option_index(select_options, auto_tema))
    sla_col = st.selectbox("SLA", select_options, index=option_index(select_options, auto_sla))
    tempo_col = st.selectbox("Tempo de resolução/atendimento", select_options, index=option_index(select_options, auto_tempo))
    prioridade_col = st.selectbox("Prioridade", select_options, index=option_index(select_options, auto_prioridade))
    canal_col = st.selectbox("Canal", select_options, index=option_index(select_options, auto_canal))
    reabertura_col = st.selectbox("Reabertura / reincidência", select_options, index=option_index(select_options, auto_reabertura))
    csat_col = st.selectbox("CSAT / satisfação", select_options, index=option_index(select_options, auto_csat))

with st.sidebar.expander("🎯 Metas e regras", expanded=False):
    default_year = st.number_input(
        "Ano padrão quando o mês estiver só no nome do arquivo",
        min_value=2020,
        max_value=2035,
        value=pd.Timestamp.today().year,
    )
    meta_sla = st.number_input("Meta de SLA cumprido (%)", min_value=0, max_value=100, value=90)
    meta_fcr = st.number_input("Meta de FCR 1h (%)", min_value=0, max_value=100, value=70)
    meta_tma_min = st.number_input("Meta de TMA em minutos", min_value=1, max_value=10000, value=20)
    sla_limite_min = st.number_input(
        "Limite para calcular SLA quando não existir coluna SLA (min)",
        min_value=1,
        max_value=10000,
        value=120,
    )

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

try:
    df = build_working_df(raw_df, mapping, int(default_year))
except Exception as exc:
    st.error(f"Erro ao tratar os dados: {exc}")
    st.stop()

if df.empty:
    st.error(
        "Não foi possível identificar o mês de referência. Use a coluna Data de abertura ou arquivos com mês no nome, como abril.xls e maio.xls."
    )
    st.stop()

summary = monthly_summary(df)
if summary.empty:
    st.error("Não foi possível gerar o resumo mensal.")
    st.stop()

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
curr_df = df[df["mes_nome"] == current["mes_nome"]]

st.caption(
    f"Arquivos carregados: **{len(uploaded_files)}** | "
    f"Linhas analisadas: **{number_br(len(df))}** | "
    f"Período: **{summary.iloc[0]['mes_nome']} a {summary.iloc[-1]['mes_nome']}**"
)


# ==========================================================
# INDICADORES PRINCIPAIS
# ==========================================================
st.markdown('<div class="section-title">Indicadores principais do mês atual</div>', unsafe_allow_html=True)

c1, c2, c3, c4, c5, c6 = st.columns(6)

prev_total = previous["total"] if previous is not None else None
prev_sla = previous["perc_sla"] if previous is not None else None
prev_tma = previous["tma_min"] if previous is not None else None
prev_fcr = previous["perc_fcr"] if previous is not None else None
prev_backlog = previous["abertos"] if previous is not None else None
prev_reopen = previous["perc_reabertura"] if previous is not None else None

with c1:
    delta, css = calc_delta(current["total"], prev_total, inverse=True)
    kpi_card("Chamados", number_br(current["total"]), f"Mês atual: {current['mes_nome']}", delta, css)
with c2:
    delta, css = calc_delta(current["perc_sla"], prev_sla, inverse=False)
    kpi_card("SLA cumprido", pct_br(current["perc_sla"]), f"Meta: {meta_sla}%", delta, css)
with c3:
    delta, css = calc_delta(current["tma_min"], prev_tma, inverse=True)
    kpi_card("TMA", min_to_human(current["tma_min"]), f"Meta: até {meta_tma_min} min", delta, css)
with c4:
    delta, css = calc_delta(current["perc_fcr"], prev_fcr, inverse=False)
    kpi_card("FCR 1h", pct_br(current["perc_fcr"]), f"Meta: {meta_fcr}%", delta, css)
with c5:
    delta, css = calc_delta(current["abertos"], prev_backlog, inverse=True)
    kpi_card("Backlog", number_br(current["abertos"]), "Chamados ainda abertos", delta, css)
with c6:
    delta, css = calc_delta(current["perc_reabertura"], prev_reopen, inverse=True)
    kpi_card("Reabertura", pct_br(current["perc_reabertura"]), "Reabertos / reincidência", delta, css)


# ==========================================================
# EVOLUÇÃO MENSAL
# ==========================================================
st.markdown('<div class="section-title">Evolução mensal e comparativos</div>', unsafe_allow_html=True)

summary_chart = summary.copy()
summary_chart["SLA %"] = summary_chart["perc_sla"].round(1)
summary_chart["FCR 1h %"] = summary_chart["perc_fcr"].round(1)
summary_chart["TMA min"] = summary_chart["tma_min"].round(1)

col_a, col_b = st.columns(2)
with col_a:
    fig = px.bar(summary_chart, x="mes_nome", y="total", text="total", title="Volume de chamados por mês")
    fig.update_traces(textposition="outside")
    st.plotly_chart(base_fig_layout(fig), use_container_width=True)

with col_b:
    fig = px.line(summary_chart, x="mes_nome", y="SLA %", markers=True, title="SLA cumprido — evolução mensal (%)")
    fig.add_hline(y=meta_sla, line_dash="dash", annotation_text=f"Meta {meta_sla}%")
    st.plotly_chart(base_fig_layout(fig), use_container_width=True)

col_c, col_d = st.columns(2)
with col_c:
    fig = px.line(summary_chart, x="mes_nome", y="FCR 1h %", markers=True, title="FCR em até 1 hora — evolução mensal (%)")
    fig.add_hline(y=meta_fcr, line_dash="dash", annotation_text=f"Meta {meta_fcr}%")
    st.plotly_chart(base_fig_layout(fig), use_container_width=True)

with col_d:
    comp = summary_chart[["mes_nome", "abertos", "finalizados", "sla_fora"]].melt(
        id_vars="mes_nome", var_name="Indicador", value_name="Quantidade"
    )
    comp["Indicador"] = comp["Indicador"].replace({
        "abertos": "Abertos/backlog",
        "finalizados": "Finalizados",
        "sla_fora": "Fora do SLA",
    })
    fig = px.bar(
        comp,
        x="mes_nome",
        y="Quantidade",
        color="Indicador",
        barmode="group",
        title="Abertos x finalizados x fora do SLA",
    )
    st.plotly_chart(base_fig_layout(fig), use_container_width=True)


# ==========================================================
# BACKLOG, CLIENTES E TEMAS
# ==========================================================
st.markdown('<div class="section-title">Backlog, clientes e temas críticos</div>', unsafe_allow_html=True)

col_e, col_f, col_g = st.columns(3)

with col_e:
    dentro = int((curr_df["aberto_calc"] & curr_df["sla_ok_calc"].fillna(False)).sum())
    fora = int((curr_df["aberto_calc"] & curr_df["sla_fora_calc"].fillna(False)).sum())
    pie_df = pd.DataFrame({"Status": ["Dentro do SLA", "Fora do SLA"], "Quantidade": [dentro, fora]})
    if pie_df["Quantidade"].sum() == 0:
        pie_df = pd.DataFrame({"Status": ["Sem backlog classificado"], "Quantidade": [1]})
    fig = px.pie(pie_df, names="Status", values="Quantidade", hole=0.55, title="Backlog — dentro x fora do SLA")
    st.plotly_chart(base_fig_layout(fig), use_container_width=True)

with col_f:
    top_clientes = curr_df["cliente_calc"].value_counts().head(8).reset_index()
    top_clientes.columns = ["Cliente", "Chamados"]
    fig = px.bar(
        top_clientes.sort_values("Chamados"),
        x="Chamados",
        y="Cliente",
        orientation="h",
        title="Top clientes do mês",
    )
    st.plotly_chart(base_fig_layout(fig), use_container_width=True)

with col_g:
    top_temas = curr_df["tema_calc"].value_counts().head(8).reset_index()
    top_temas.columns = ["Tema", "Chamados"]
    fig = px.bar(
        top_temas.sort_values("Chamados"),
        x="Chamados",
        y="Tema",
        orientation="h",
        title="Top temas/motivos do mês",
    )
    st.plotly_chart(base_fig_layout(fig), use_container_width=True)


# ==========================================================
# ANÁLISE EXECUTIVA AUTOMÁTICA
# ==========================================================
st.markdown('<div class="section-title">Leitura executiva automática</div>', unsafe_allow_html=True)

alerts = []
success = []
recommendations = []

if current["perc_sla"] < meta_sla:
    alerts.append(f"SLA abaixo da meta: {pct_br(current['perc_sla'])}, meta de {meta_sla}%.")
    recommendations.append("Priorizar chamados fora do SLA e revisar gargalos por cliente, tema e fila.")
else:
    success.append(f"SLA dentro/acima da meta: {pct_br(current['perc_sla'])}.")

if current["perc_fcr"] < meta_fcr:
    alerts.append(f"FCR 1h abaixo da meta: {pct_br(current['perc_fcr'])}, meta de {meta_fcr}%.")
    recommendations.append("Criar base de conhecimento, scripts e automações para resolver mais chamados no primeiro contato.")
else:
    success.append(f"FCR 1h dentro/acima da meta: {pct_br(current['perc_fcr'])}.")

if pd.notna(current["tma_min"]) and current["tma_min"] > meta_tma_min:
    alerts.append(f"TMA acima da meta: {min_to_human(current['tma_min'])}, meta de até {meta_tma_min} min.")
    recommendations.append("Analisar os motivos mais recorrentes e automatizar atendimentos repetitivos.")
else:
    success.append(f"TMA dentro da meta ou sem dado suficiente para alerta: {min_to_human(current['tma_min'])}.")

if current["abertos"] > 0:
    alerts.append(f"Backlog atual: {number_br(current['abertos'])} chamado(s) aberto(s).")
    recommendations.append("Separar backlog crítico dos chamados de baixa prioridade e acompanhar diariamente.")

if previous is not None:
    if current["total"] > previous["total"]:
        alerts.append(f"Volume aumentou em relação ao mês anterior: {number_br(previous['total'])} → {number_br(current['total'])}.")
    elif current["total"] < previous["total"]:
        success.append(f"Volume reduziu em relação ao mês anterior: {number_br(previous['total'])} → {number_br(current['total'])}.")

if not alerts:
    alerts.append("Nenhum risco crítico identificado com as regras atuais.")
if not success:
    success.append("Ainda não há evolução positiva suficiente para destacar com os dados atuais.")
if not recommendations:
    recommendations.append("Manter acompanhamento mensal e revisar metas conforme maturidade do suporte.")

col_h, col_i, col_j = st.columns(3)
with col_h:
    st.markdown(
        "<div class='alert-box'><b>Riscos / pontos de atenção</b><br>" +
        "<br>".join([f"• {a}" for a in alerts[:6]]) +
        "</div>",
        unsafe_allow_html=True,
    )
with col_i:
    st.markdown(
        "<div class='success-box'><b>Evoluções positivas</b><br>" +
        "<br>".join([f"• {s}" for s in success[:6]]) +
        "</div>",
        unsafe_allow_html=True,
    )
with col_j:
    unique_recs = list(dict.fromkeys(recommendations))
    st.markdown(
        "<div class='info-box'><b>Recomendações</b><br>" +
        "<br>".join([f"• {r}" for r in unique_recs[:6]]) +
        "</div>",
        unsafe_allow_html=True,
    )


# ==========================================================
# TABELAS DE APOIO
# ==========================================================
st.markdown('<div class="section-title">Tabelas de apoio</div>', unsafe_allow_html=True)

summary_view = summary.copy()
summary_view = summary_view[[
    "mes_nome", "total", "abertos", "finalizados", "sla_ok", "sla_fora",
    "perc_sla", "fcr_1h", "perc_fcr", "tma_min", "reabertos",
    "perc_reabertura", "csat",
]]
summary_view.columns = [
    "Mês", "Total", "Abertos", "Finalizados", "SLA cumprido", "Fora do SLA",
    "% SLA", "FCR 1h", "% FCR 1h", "TMA (min)", "Reabertos",
    "% Reabertura", "CSAT",
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
