from __future__ import annotations

from html import escape
from io import BytesIO
from pathlib import Path
import os

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
RESUMO_XLSX = BASE_DIR / "Resumo Diario de Vendas - Fardin 2026.xlsx"
AVALIACOES_XLSX = BASE_DIR / "Avaliacoes de Desempenho Comercial - Doces Fardin 2026.xlsx"
LOGO_PATH = BASE_DIR / "logo fardin.jpg"
ALFA_LOGO_PATH = BASE_DIR / "Logo Alfa.jpg"
_SENHA_COMPLETA = "zampa255"
_SENHA_DEMO = "zampa"
DEMO_ALLOWED_MONTHS = {(2026, 5), (2026, 6), (2026, 7)}

# Fallback para os nomes originais com acento, usados nos arquivos recebidos.
if not RESUMO_XLSX.exists():
    RESUMO_XLSX = BASE_DIR / "Resumo Diário de Vendas - Fardin 2026.xlsx"
if not AVALIACOES_XLSX.exists():
    AVALIACOES_XLSX = BASE_DIR / "Avaliações de Desempenho Comercial - Doces Fardin 2026.xlsx"


WorkbookSource = str | Path | bytes


def workbook_source(source: WorkbookSource):
    return BytesIO(source) if isinstance(source, bytes) else source


def excel_file(source: WorkbookSource) -> pd.ExcelFile:
    return pd.ExcelFile(workbook_source(source))


def read_excel_workbook(source: WorkbookSource, **kwargs) -> pd.DataFrame:
    return pd.read_excel(workbook_source(source), **kwargs)


def source_name(source: WorkbookSource, uploaded_name: str | None, default_path: Path) -> str:
    if uploaded_name:
        return uploaded_name
    return default_path.name

def is_streamlit_cloud() -> bool:
    return bool(os.environ.get("STREAMLIT_SHARING") or os.environ.get("STREAMLIT_CLOUD"))

MONTH_SHEETS = ("Fev26", "Mar26", "Abr26", "Mai26", "Jun26", "Jul26")
MONTH_ORDER = {
    "Janeiro": 1,
    "Fevereiro": 2,
    "Março": 3,
    "Abril": 4,
    "Maio": 5,
    "Junho": 6,
    "Julho": 7,
    "Agosto": 8,
    "Setembro": 9,
    "Outubro": 10,
    "Novembro": 11,
    "Dezembro": 12,
}


st.set_page_config(
    page_title="Dashboard Comercial - Doces Fardin",
    page_icon="📊",
    layout="wide",
)



def br_date_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.strftime("%d/%m/%Y").fillna("")


def br_month_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.strftime("%m/%Y").fillna("")
def money(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return "R$ 0,00"
    return "R$ " + f"{float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def number(value: float | int | None, decimals: int = 0) -> str:
    if value is None or pd.isna(value):
        value = 0
    text = f"{float(value):,.{decimals}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        return "0,0%"
    return f"{float(value) * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def to_num(value) -> float:
    if value is None or pd.isna(value) or str(value).strip() in {"", "-"}:
        return 0.0
    return float(value)


def first_row_value(df: pd.DataFrame, label: str, offset: int = 1):
    mask = df.apply(lambda row: row.astype(str).str.strip().eq(label).any(), axis=1)
    if not mask.any():
        return None
    r = int(np.where(mask)[0][0])
    c = int(np.where(df.iloc[r].astype(str).str.strip().eq(label))[0][0])
    if c + offset >= df.shape[1]:
        return None
    return df.iat[r, c + offset]


@st.cache_data(show_spinner=False)
def load_monthly_summary(path: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    xls = excel_file(path)
    daily_frames: list[pd.DataFrame] = []
    seller_frames: list[pd.DataFrame] = []
    meta_rows: list[dict] = []

    for sheet in [s for s in MONTH_SHEETS if s in xls.sheet_names]:
        raw = read_excel_workbook(path, sheet_name=sheet, header=None)
        ref_month = pd.to_datetime(first_row_value(raw, "Mês de Ref.:")).normalize()
        meta_mes = to_num(first_row_value(raw, "Meta Mês:"))
        meta_dia = to_num(first_row_value(raw, "Meta Dia:"))
        dias_uteis = to_num(first_row_value(raw, "Dias Úteis Mês"))

        row_labels = raw.iloc[6].fillna("").astype(str).str.strip()
        seller_start = None
        for idx, label in enumerate(row_labels):
            if label in {"Entrada", "Fardin"}:
                seller_start = idx
                break
        seller_cols: list[tuple[int, str]] = []
        if seller_start is not None:
            for col in range(seller_start, raw.shape[1]):
                label = str(row_labels.iloc[col]).strip()
                if not label:
                    if seller_cols:
                        break
                    continue
                if label == "Realizado" and seller_cols:
                    break
                seller_cols.append((col, label))

        rows = []
        seller_rows = []
        for ridx in range(9, raw.shape[0]):
            data = pd.to_datetime(raw.iat[ridx, 1], errors="coerce")
            if pd.isna(data):
                continue
            if pd.notna(ref_month) and (data.month != ref_month.month or data.year != ref_month.year):
                continue

            faturamento_dia = to_num(raw.iat[ridx, 7])
            faturamento_acum = to_num(raw.iat[ridx, 8])
            qtd_pedidos = to_num(raw.iat[ridx, 10])
            entrada_total = sum(to_num(raw.iat[ridx, c]) for c, _ in seller_cols)

            rows.append(
                {
                    "aba": sheet,
                    "mes_ref": ref_month,
                    "data": data,
                    "mes": data.strftime("%b/%y"),
                    "dia_semana": str(raw.iat[ridx, 2]).strip(),
                    "dia_util": int(to_num(raw.iat[ridx, 4])),
                    "dia_util_mes": int(to_num(raw.iat[ridx, 5])),
                    "resultado_projetado": to_num(raw.iat[ridx, 6]),
                    "faturamento_realizado": faturamento_dia,
                    "faturamento_acumulado": faturamento_acum,
                    "atingimento_dia": to_num(raw.iat[ridx, 9]),
                    "qtd_pedidos": qtd_pedidos,
                    "entrada_pv": entrada_total,
                    "ticket_medio": to_num(raw.iat[ridx, 21]) if raw.shape[1] > 21 else 0,
                    "backlog_pedidos": to_num(raw.iat[ridx, 22]) if raw.shape[1] > 22 else 0,
                    "ajuste": to_num(raw.iat[ridx, 24]) if raw.shape[1] > 24 else 0,
                    "meta_mes": meta_mes,
                    "meta_dia": meta_dia,
                    "dias_uteis_mes": dias_uteis,
                }
            )

            for col, seller in seller_cols:
                value = to_num(raw.iat[ridx, col])
                seller_rows.append(
                    {
                        "aba": sheet,
                        "mes_ref": ref_month,
                        "data": data,
                        "mes": data.strftime("%b/%y"),
                        "vendedor": seller,
                        "entrada_pv": value,
                        "meta_vendedor_mes": to_num(raw.iat[7, col]) if raw.shape[0] > 7 else 0,
                    }
                )

        meta_rows.append(
            {
                "aba": sheet,
                "mes_ref": ref_month,
                "mes": ref_month.strftime("%b/%y") if pd.notna(ref_month) else sheet,
                "meta_mes": meta_mes,
                "meta_dia": meta_dia,
                "dias_uteis_mes": dias_uteis,
                "vendedores": ", ".join(label for _, label in seller_cols),
            }
        )

        daily_frames.append(pd.DataFrame(rows))
        seller_frames.append(pd.DataFrame(seller_rows))

    daily = pd.concat(daily_frames, ignore_index=True) if daily_frames else pd.DataFrame()
    sellers = pd.concat(seller_frames, ignore_index=True) if seller_frames else pd.DataFrame()
    metas = pd.DataFrame(meta_rows)
    return daily, sellers, metas


@st.cache_data(show_spinner=False)
def load_meta_global(path: str) -> pd.DataFrame:
    raw = read_excel_workbook(path, sheet_name="Meta x Real de Vendas  Global", header=None)
    rows = []
    for idx in range(6, min(raw.shape[0], 18)):
        mes = raw.iat[idx, 1]
        if pd.isna(mes):
            continue
        rows.append(
            {
                "mes": str(mes).strip(),
                "ordem": MONTH_ORDER.get(str(mes).strip(), idx),
                "meta": to_num(raw.iat[idx, 2]),
                "total": to_num(raw.iat[idx, 3]),
                "fardin": to_num(raw.iat[idx, 4]),
                "privaty_label": to_num(raw.iat[idx, 5]),
                "terceiros": to_num(raw.iat[idx, 6]),
                "atingimento": to_num(raw.iat[idx, 7]),
                "venda_acumulada": to_num(raw.iat[idx, 8]),
                "falta_meta": to_num(raw.iat[idx, 9]),
            }
        )
    return pd.DataFrame(rows).sort_values("ordem")


@st.cache_data(show_spinner=False)
def load_meta_fardin(path: str) -> pd.DataFrame:
    raw = read_excel_workbook(path, sheet_name="Meta x Real de Vendas  FARDIN", header=None)
    rows = []
    for idx in range(6, min(raw.shape[0], 18)):
        mes = raw.iat[idx, 1]
        if pd.isna(mes):
            continue
        rows.append(
            {
                "mes": str(mes).strip(),
                "ordem": MONTH_ORDER.get(str(mes).strip(), idx),
                "representantes_meta": to_num(raw.iat[idx, 10]),
                "representantes_real": to_num(raw.iat[idx, 11]),
                "doces_fardin_meta": to_num(raw.iat[idx, 12]),
                "doces_fardin_real": to_num(raw.iat[idx, 13]),
                "wilson_meta": to_num(raw.iat[idx, 14]),
                "wilson_real": to_num(raw.iat[idx, 15]),
                "total_meta": to_num(raw.iat[idx, 16]),
                "total_real": to_num(raw.iat[idx, 17]),
                "atingimento": to_num(raw.iat[idx, 18]),
                "venda_acumulada": to_num(raw.iat[idx, 19]),
                "falta_meta": to_num(raw.iat[idx, 20]),
            }
        )
    return pd.DataFrame(rows).sort_values("ordem")


@st.cache_data(show_spinner=False)
def load_evolution(path: str, sheet: str) -> pd.DataFrame:
    raw = read_excel_workbook(path, sheet_name=sheet, header=None)
    rows = []
    for idx in range(5, min(raw.shape[0], 17)):
        base_date = pd.to_datetime(raw.iat[idx, 1], errors="coerce")
        real_date = pd.to_datetime(raw.iat[idx, 3], errors="coerce")
        if pd.isna(base_date):
            continue
        rows.append(
            {
                "mes": base_date.strftime("%b"),
                "base_data": base_date,
                "base_vendas": to_num(raw.iat[idx, 2]),
                "real_data": real_date,
                "real_vendas": to_num(raw.iat[idx, 4]),
                "evolucao_vendas": to_num(raw.iat[idx, 5]),
                "ganho_acumulado": to_num(raw.iat[idx, 6]),
                "media_mensal": to_num(raw.iat[idx, 7]),
                "anualizado": to_num(raw.iat[idx, 8]),
                "crescimento": to_num(raw.iat[idx, 10]) if raw.shape[1] > 10 else 0,
            }
        )
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def load_history_pf(path: str) -> pd.DataFrame:
    raw = read_excel_workbook(path, sheet_name="Histórico de Fatu. FARDIN P.F.", header=None)
    months = list(raw.iloc[5, 4:16])
    rows = []
    current_family = None
    for idx in range(10, raw.shape[0]):
        first = raw.iat[idx, 2]
        metric = raw.iat[idx, 3] if raw.shape[1] > 3 else None
        if pd.notna(first) and str(first).strip() not in {"Real", "% de Atingimento", "% de Participação"}:
            current_family = str(first).strip()
            metric = raw.iat[idx, 3]
        elif pd.notna(first):
            metric = first
        if not current_family or pd.isna(metric):
            continue
        for pos, month in enumerate(months, start=4):
            if pd.isna(month):
                continue
            rows.append(
                {
                    "familia": current_family,
                    "indicador": str(metric).strip(),
                    "mes": str(month).strip(),
                    "valor": to_num(raw.iat[idx, pos]),
                }
            )
    return pd.DataFrame(rows)


FAMILY_MONTH_ORDER = {
    "Jan": 1,
    "Fev": 2,
    "Mar": 3,
    "Abri": 4,
    "Abr": 4,
    "Mai": 5,
    "Jun": 6,
    "Jul": 7,
    "Ago": 8,
    "Set": 9,
    "Out": 10,
    "Nov": 11,
    "Dez": 12,
}


def family_month_ref(label: str) -> pd.Timestamp | pd.NaT:
    month = FAMILY_MONTH_ORDER.get(str(label).strip())
    if not month:
        return pd.NaT
    return pd.Timestamp(year=2026, month=month, day=1)


@st.cache_data(show_spinner=False)
def load_family_performance(path: WorkbookSource) -> pd.DataFrame:
    history = load_history_pf(path)
    if history.empty:
        return pd.DataFrame(columns=["familia", "mes_ref", "mes", "meta_familia", "realizado_familia", "atingimento_familia"])

    rows = []
    for (family, month_label), group in history.groupby(["familia", "mes"], dropna=True):
        month_ref = family_month_ref(month_label)
        if pd.isna(month_ref):
            continue
        meta = group.loc[group["indicador"].eq("Meta"), "valor"].sum()
        realised = group.loc[group["indicador"].eq("Real"), "valor"].sum()
        if meta == 0 and realised == 0:
            continue
        rows.append(
            {
                "familia": str(family).strip(),
                "mes_ref": month_ref,
                "mes": month_ref.strftime("%b/%y"),
                "meta_familia": float(meta),
                "realizado_familia": float(realised),
                "atingimento_familia": float(realised / meta) if meta else 0.0,
            }
        )
    return pd.DataFrame(rows)

@st.cache_data(show_spinner=False)
def load_history_pr(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = read_excel_workbook(path, sheet_name="Histórico de Fatu. FARDIN P.R.", header=None)
    months = list(raw.iloc[14, 4:16])

    geral = []
    for idx in range(6, 12):
        indicador = raw.iat[idx, 2]
        if pd.isna(indicador):
            continue
        for pos, month in enumerate(months, start=4):
            if pd.isna(month):
                continue
            geral.append(
                {
                    "indicador": str(indicador).strip(),
                    "mes": str(month).strip(),
                    "valor": to_num(raw.iat[idx, pos]),
                }
            )

    reps = []
    current_rep = None
    for idx in range(15, raw.shape[0]):
        first = raw.iat[idx, 2]
        metric = raw.iat[idx, 3] if raw.shape[1] > 3 else None
        if pd.notna(first) and str(first).strip() not in {
            "Venda (PV)",
            "Atingimento",
            "Faturamento (NF)",
            "Nº Clientes Carteira",
            "Nº De Vendas",
        }:
            current_rep = str(first).strip()
            metric = raw.iat[idx, 3]
        elif pd.notna(first):
            metric = first
        if not current_rep or pd.isna(metric):
            continue
        for pos, month in enumerate(months, start=4):
            if pd.isna(month):
                continue
            reps.append(
                {
                    "representante": current_rep,
                    "indicador": str(metric).strip(),
                    "mes": str(month).strip(),
                    "valor": to_num(raw.iat[idx, pos]),
                }
            )

    return pd.DataFrame(geral), pd.DataFrame(reps)


@st.cache_data(show_spinner=False)
def load_budget(path: str) -> pd.DataFrame:
    raw = read_excel_workbook(path, sheet_name="Orçamento Comercial", header=None)
    rows = []
    for idx in range(5, raw.shape[0]):
        mes = raw.iat[idx, 1]
        if pd.isna(mes):
            continue
        rows.append(
            {
                "mes": mes if isinstance(mes, str) else pd.to_datetime(mes, errors="coerce"),
                "viagens_planejado": to_num(raw.iat[idx, 2]),
                "viagens_realizado": to_num(raw.iat[idx, 3]),
                "viagens_pct": to_num(raw.iat[idx, 4]),
                "pdv_planejado": to_num(raw.iat[idx, 5]),
                "pdv_realizado": to_num(raw.iat[idx, 6]),
                "pdv_pct": to_num(raw.iat[idx, 7]),
                "eventos_planejado": to_num(raw.iat[idx, 8]),
                "eventos_realizado": to_num(raw.iat[idx, 9]),
                "eventos_pct": to_num(raw.iat[idx, 10]),
            }
        )
    return pd.DataFrame(rows)



def find_sheet(path: str, text: str) -> str:
    xls = excel_file(path)
    text_norm = text.casefold().strip()
    for sheet in xls.sheet_names:
        if text_norm in sheet.casefold().strip():
            return sheet
    raise ValueError(f"Aba contendo '{text}' nao encontrada em {path}")


@st.cache_data(show_spinner=False)
def load_agenda_impact(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    sheet = find_sheet(path, "Agenda")
    raw = read_excel_workbook(path, sheet_name=sheet, header=None)
    month_names = {"JANEIRO", "FEVEREIRO", "MARÇO", "MARCO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"}
    weekdays = {"do", "se", "te", "qu", "sá", "sa"}
    points = []
    for ridx in range(raw.shape[0]):
        for cidx in range(raw.shape[1]):
            value = raw.iat[ridx, cidx]
            if pd.isna(value):
                continue
            item = str(value).strip()
            if not item or item.upper() in month_names or item.casefold() in weekdays:
                continue
            if item.isdigit() or item == "2026" or "CALEND" in item.upper():
                continue
            if "Pontos Importantes" in item:
                continue
            if ridx >= 27:
                points.append({"ponto_de_impacto": item})
    points_df = pd.DataFrame(points).drop_duplicates().reset_index(drop=True)
    calendar_df = raw.dropna(how="all").dropna(axis=1, how="all")
    return points_df, calendar_df


@st.cache_data(show_spinner=False)
def load_per_capita(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = read_excel_workbook(path, sheet_name=find_sheet(path, "per Capita"), header=None)
    base_rows = []
    for idx in range(2, 15):
        mes = raw.iat[idx, 5]
        if pd.isna(mes):
            continue
        if str(mes).strip().casefold() == "média":
            label = "Media 2023"
            data = pd.NaT
        else:
            label = "2023"
            data = pd.to_datetime(mes, errors="coerce")
        base_rows.append({"periodo": label, "data": data, "mes": "Media" if pd.isna(data) else data.strftime("%b/%y"), "vendas": to_num(raw.iat[idx, 6]), "funcionarios": to_num(raw.iat[idx, 7]), "per_capita": to_num(raw.iat[idx, 8]), "ganho": to_num(raw.iat[idx, 11]) if raw.shape[1] > 11 else 0, "acumulado": 0, "media_mensal": 0, "anualizado": 0})
    real_rows = []
    for idx in range(17, raw.shape[0]):
        data = pd.to_datetime(raw.iat[idx, 1], errors="coerce")
        if pd.isna(data):
            continue
        real_rows.append({"periodo": "2024", "data": data, "mes": data.strftime("%b/%y"), "vendas": to_num(raw.iat[idx, 2]), "funcionarios": to_num(raw.iat[idx, 3]), "per_capita": to_num(raw.iat[idx, 4]), "ganho": to_num(raw.iat[idx, 5]), "acumulado": to_num(raw.iat[idx, 6]), "media_mensal": to_num(raw.iat[idx, 7]), "anualizado": to_num(raw.iat[idx, 8])})
    detail = pd.concat([pd.DataFrame(base_rows), pd.DataFrame(real_rows)], ignore_index=True)
    detail = detail[(detail["vendas"] > 0) | (detail["per_capita"] > 0)].copy()
    summary = detail.groupby("periodo", as_index=False).agg(vendas=("vendas", "mean"), funcionarios=("funcionarios", "mean"), per_capita=("per_capita", "mean"), ganho=("ganho", "sum")).sort_values("periodo")
    return detail, summary


@st.cache_data(show_spinner=False)
def load_graficos_raw(path: str) -> pd.DataFrame:
    raw = read_excel_workbook(path, sheet_name=find_sheet(path, "Gráficos"), header=None)
    return raw.dropna(how="all").dropna(axis=1, how="all")
def is_demo_access() -> bool:
    return st.session_state.get("_access_mode") == "demo"

def require_login() -> None:
    if st.session_state.get("_autenticado"):
        st.session_state.setdefault("_access_mode", "full")
        return

    st.markdown(
        """
        <style>
        .stApp {
            background: #0f172a;
            color: #e5e7eb;
        }
        [data-testid="stSidebar"],
        [data-testid="stHeader"] {
            display: none;
        }
        .block-container {
            max-width: 980px;
            padding-top: 4rem;
        }
        div[data-testid="stForm"] {
            background: #111827;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 18px 18px 8px;
        }
        div[data-testid="stImage"] img {
            border-radius: 8px;
            background: #ffffff;
            padding: 10px;
            box-shadow: 0 10px 28px rgba(0, 0, 0, .28);
        }
        h2, label, p, span {
            color: #e5e7eb;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col_left, col_center, col_right = st.columns([1, 1, 1])
    with col_center:
        st.markdown("<div style='margin-top:40px'></div>", unsafe_allow_html=True)
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), use_container_width=True)
        st.markdown(
            "<h2 style='text-align:center;margin-top:14px'>Area Restrita - Fardin</h2>",
            unsafe_allow_html=True,
        )
        with st.form("_login_form"):
            senha = st.text_input("Senha de acesso", type="password")
            submitted = st.form_submit_button("Entrar", use_container_width=True, type="primary")
        if submitted:
            if senha == _SENHA_COMPLETA:
                st.session_state["_autenticado"] = True
                st.session_state["_access_mode"] = "full"
                st.rerun()
            elif senha == _SENHA_DEMO:
                st.session_state["_autenticado"] = True
                st.session_state["_access_mode"] = "demo"
                st.rerun()
            else:
                st.error("Senha incorreta.")
    st.stop()

def select_theme() -> str:
    with st.sidebar:
        st.header("Aparência")
        choice = st.radio(
            "Tema",
            ["Escuro", "Claro"],
            index=0,
            horizontal=True,
            key="app_theme_choice",
        )
        access_label = "Demonstração" if is_demo_access() else "Completo"
        st.caption(f"Acesso: {access_label}")
        if st.button("Sair", use_container_width=True):
            st.session_state.pop("_autenticado", None)
            st.session_state.pop("_access_mode", None)
            st.rerun()
        st.divider()
    return "dark" if choice == "Escuro" else "light"

def select_data_sources() -> tuple[WorkbookSource | None, WorkbookSource | None, str, str]:
    require_upload = is_streamlit_cloud()
    with st.sidebar:
        st.markdown("### Arquivos de dados")
        st.caption("Envie as planilhas atualizadas da Fardin para carregar o painel.")

        st.markdown("<div class='upload-title'>Resumo diario de vendas</div>", unsafe_allow_html=True)
        resumo_upload = st.file_uploader(
            "Resumo diario de vendas",
            type=["xlsx"],
            key="resumo_xlsx_upload",
            label_visibility="collapsed",
        )
        resumo_source: WorkbookSource | None = resumo_upload.getvalue() if resumo_upload else (None if require_upload else RESUMO_XLSX)
        resumo_name = resumo_upload.name if resumo_upload else ("Aguardando upload" if require_upload else RESUMO_XLSX.name)
        resumo_status = "Selecionado" if resumo_upload else ("Obrigatorio" if require_upload else "Padrao")
        st.markdown(
            f"<div class='file-status'><span>{resumo_status}</span><strong>{escape(resumo_name)}</strong></div>",
            unsafe_allow_html=True,
        )

        st.markdown("<div class='upload-title upload-title-spaced'>Avaliacoes de desempenho</div>", unsafe_allow_html=True)
        avaliacoes_upload = st.file_uploader(
            "Avaliacoes de desempenho",
            type=["xlsx"],
            key="avaliacoes_xlsx_upload",
            label_visibility="collapsed",
        )
        avaliacoes_source: WorkbookSource | None = avaliacoes_upload.getvalue() if avaliacoes_upload else (None if require_upload else AVALIACOES_XLSX)
        avaliacoes_name = avaliacoes_upload.name if avaliacoes_upload else ("Aguardando upload" if require_upload else AVALIACOES_XLSX.name)
        avaliacoes_status = "Selecionado" if avaliacoes_upload else ("Obrigatorio" if require_upload else "Padrao")
        st.markdown(
            f"<div class='file-status'><span>{avaliacoes_status}</span><strong>{escape(avaliacoes_name)}</strong></div>",
            unsafe_allow_html=True,
        )
        st.divider()

    return resumo_source, avaliacoes_source, resumo_name, avaliacoes_name

def style_app(theme: str) -> None:
    if theme == "light":
        colors = {
            "app_bg": "#f3f6fa",
            "panel_bg": "#ffffff",
            "text": "#111827",
            "muted": "#5b667a",
            "metric_label": "#374151",
            "border": "#d8dee8",
            "sidebar_bg": "#ffffff",
            "input_bg": "#f8fafc",
            "hero_start": "#172554",
            "hero_mid": "#1e3a5f",
            "hero_end": "#166534",
            "hero_text": "#ffffff",
            "hero_muted": "#dbeafe",
            "delta_text": "#047857",
            "delta_bg": "#dcfce7",
        }
    else:
        colors = {
            "app_bg": "#0f172a",
            "panel_bg": "#111827",
            "text": "#e5e7eb",
            "muted": "#94a3b8",
            "metric_label": "#cbd5e1",
            "border": "#334155",
            "sidebar_bg": "#020617",
            "input_bg": "#1e293b",
            "hero_start": "#020617",
            "hero_mid": "#172554",
            "hero_end": "#14532d",
            "hero_text": "#f8fafc",
            "hero_muted": "#bfdbfe",
            "delta_text": "#86efac",
            "delta_bg": "rgba(22, 101, 52, .42)",
        }

    st.markdown(
        f"""
        <style>
        .stApp {{
            background: {colors["app_bg"]};
            color: {colors["text"]};
        }}
        .block-container {{
            max-width: 1480px;
            padding-top: 1.2rem;
            padding-right: 2rem;
            padding-left: 2rem;
            padding-bottom: 2rem;
        }}
        @media (min-width: 1200px) {{
            .block-container {{
                max-width: 1540px;
            }}
        }}
        [data-testid="stSidebar"] {{
            background: {colors["sidebar_bg"]};
            border-right: 1px solid {colors["border"]};
        }}
        [data-testid="stSidebar"] * {{
            color: {colors["text"]};
        }}
        [data-testid="stHeader"] {{
            background: transparent;
        }}
        [data-testid="stTabs"] button {{
            color: {colors["muted"]};
        }}
        [data-testid="stTabs"] button[aria-selected="true"] {{
            color: {colors["text"]};
        }}
        div[data-testid="stMetric"], div[data-testid="stDataFrame"] {{
            background: {colors["panel_bg"]};
            border: 1px solid {colors["border"]};
            border-radius: 8px;
            padding: 10px 12px;
        }}
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        div[data-baseweb="radio"] {{
            background: {colors["input_bg"]};
            color: {colors["text"]};
        }}
        .upload-title {{
            margin: 14px 0 6px;
            color: {colors["text"]};
            font-size: .88rem;
            font-weight: 700;
        }}
        .upload-title-spaced {{
            margin-top: 18px;
        }}
        .file-status {{
            margin-top: 6px;
            padding: 8px 10px;
            border: 1px solid {colors["border"]};
            border-radius: 8px;
            background: {colors["panel_bg"]};
        }}
        .file-status span {{
            display: block;
            margin-bottom: 3px;
            color: {colors["muted"]};
            font-size: .72rem;
            font-weight: 700;
            text-transform: uppercase;
        }}
        .file-status strong {{
            display: block;
            color: {colors["text"]};
            font-size: .78rem;
            line-height: 1.25;
            overflow-wrap: anywhere;
        }}
        [data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
            min-height: 0;
            padding: 9px;
            border: 1px solid {colors["border"]};
            border-radius: 8px;
            background: {colors["input_bg"]};
        }}
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 46px;
            padding: 6px;
        }}
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] svg,
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] p,
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] small,
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] span,
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] [data-testid="stFileUploaderDropzoneInstructions"] {{
            display: none !important;
        }}
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {{
            width: 100%;
            min-height: 34px;
            border-radius: 6px;
            font-size: 0 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button::after {{
            content: "Selecionar planilha";
            font-size: .82rem;
            font-weight: 700;
        }}
        h1, h2, h3, h4, h5, h6, p, label, span {{
            color: inherit;
        }}
        [data-testid="stMetricValue"] {{
            color: {colors["text"]};
            font-size: clamp(1.25rem, 1.5vw, 1.65rem);
            white-space: normal;
            overflow-wrap: anywhere;
        }}
        [data-testid="stMetricDelta"] {{
            color: {colors["delta_text"]};
            background: {colors["delta_bg"]};
            border-radius: 999px;
            padding: 2px 8px;
            width: fit-content;
            font-weight: 700;
        }}
        [data-testid="stMetricDelta"] svg {{
            fill: {colors["delta_text"]};
        }}
        [data-testid="stMetric"] {{box-shadow: 0 1px 3px rgba(15, 23, 42, .08);}}
        [data-testid="stSidebar"] .stButton button {{
            border: 1px solid {colors["border"]};
            background: {colors["input_bg"]};
            color: {colors["text"]};
        }}
        [data-testid="stMetricLabel"] {{font-weight: 700; color: {colors["metric_label"]};}}
        .hero {{
            padding: 18px 22px;
            border-radius: 8px;
            background: linear-gradient(135deg, {colors["hero_start"]} 0%, {colors["hero_mid"]} 62%, {colors["hero_end"]} 100%);
            color: {colors["hero_text"]};
            margin-bottom: 16px;
        }}
        .hero h1 {{font-size: 1.7rem; margin: 0 0 4px 0; color: {colors["hero_text"]};}}
        .hero p {{margin: 0; color: {colors["hero_muted"]};}}
        div[data-testid="stImage"] img {{
            border-radius: 6px;
            background: #ffffff;
            padding: 8px;
            box-shadow: 0 8px 22px rgba(0, 0, 0, .18);
        }}
        @media (max-width: 640px) {{
            div[data-testid="stImage"] img {{
                max-width: 96px;
            }}
            .hero h1 {{
                font-size: 1.35rem;
            }}
        }}
        .file-required-panel {{
            display: flex;
            flex-direction: column;
            gap: 7px;
            margin: 16px 0 24px;
            padding: 18px 20px;
            border: 1px solid {colors["border"]};
            border-radius: 8px;
            background: {colors["panel_bg"]};
        }}
        .file-required-panel strong {{
            color: {colors["text"]};
            font-size: 1rem;
        }}
        .file-required-panel span {{
            color: {colors["muted"]};
            font-size: .9rem;
        }}
        .app-footer {{
            margin-top: 28px;
            padding: 14px 18px;
            border-top: 1px solid {colors["border"]};
            color: {colors["muted"]};
        }}
        .footer-copy {{
            display: flex;
            flex-direction: column;
            gap: 2px;
            line-height: 1.25;
        }}
        .footer-copy strong {{
            color: {colors["text"]};
            font-size: .9rem;
        }}
        .footer-copy span {{
            color: {colors["muted"]};
            font-size: .82rem;
        }}
        .app-footer div[data-testid="stImage"] img {{
            max-width: 118px;
            margin-left: auto;
            box-shadow: none;
        }}
        @media (max-width: 640px) {{
            .app-footer {{
                padding: 12px 0;
            }}
        }}
        .section-note {{
            color: {colors["muted"]};
            font-size: .9rem;
            margin-top: -6px;
            margin-bottom: 12px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def metric_card(label: str, value: str, delta: str | None = None) -> None:
    st.metric(label, value, delta=delta)


def apply_chart_theme(fig):
    if st.session_state.get("_theme") == "light":
        fig.update_layout(
            template="plotly_white",
            paper_bgcolor="#ffffff",
            plot_bgcolor="#ffffff",
            font=dict(color="#111827"),
            title_font=dict(color="#111827"),
            legend=dict(font=dict(color="#111827")),
            margin=dict(l=50, r=24, t=58, b=48),
        )
        fig.update_xaxes(
            color="#111827",
            title_font=dict(color="#111827"),
            tickfont=dict(color="#111827"),
            gridcolor="#e5e7eb",
            zerolinecolor="#d1d5db",
            linecolor="#cbd5e1",
        )
        fig.update_yaxes(
            color="#111827",
            title_font=dict(color="#111827"),
            tickfont=dict(color="#111827"),
            gridcolor="#e5e7eb",
            zerolinecolor="#d1d5db",
            linecolor="#cbd5e1",
        )
    return fig


def plot_chart(fig, **kwargs) -> None:
    kwargs.setdefault("use_container_width", True)
    st.plotly_chart(apply_chart_theme(fig), **kwargs)


def format_table_cell(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (float, np.floating)):
        return f"{float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if isinstance(value, (int, np.integer)):
        return f"{int(value):,}".replace(",", ".")
    return escape(str(value))


def render_table(df: pd.DataFrame, **kwargs) -> None:
    if st.session_state.get("_theme") != "light":
        st.dataframe(df, **kwargs)
        return

    table_df = df.copy()
    rows = []
    for _, row in table_df.iterrows():
        cells = "".join(f"<td>{format_table_cell(value)}</td>" for value in row)
        rows.append(f"<tr>{cells}</tr>")
    headers = "".join(f"<th>{escape(str(col))}</th>" for col in table_df.columns)
    st.markdown(
        f"""
        <div class="light-table-wrap">
          <table class="light-table">
            <thead><tr>{headers}</tr></thead>
            <tbody>{''.join(rows)}</tbody>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )

def seller_goal_table(sellers: pd.DataFrame) -> pd.DataFrame:
    if sellers.empty:
        return pd.DataFrame(columns=["Vendedor/canal", "Realizado PV", "Meta"])

    seller_actual = sellers.groupby("vendedor", as_index=False)["entrada_pv"].sum()
    if "meta_vendedor_mes" in sellers.columns:
        seller_defaults = (
            sellers.drop_duplicates(["mes_ref", "vendedor"])
            .groupby("vendedor", as_index=False)["meta_vendedor_mes"]
            .sum()
        )
    else:
        seller_defaults = pd.DataFrame({"vendedor": seller_actual["vendedor"], "meta_vendedor_mes": 0.0})

    goal_table = seller_actual.merge(seller_defaults, on="vendedor", how="left").fillna(0)
    goal_table = goal_table.rename(
        columns={
            "vendedor": "Vendedor/canal",
            "entrada_pv": "Realizado PV",
            "meta_vendedor_mes": "Meta",
        }
    )
    return goal_table.sort_values("Realizado PV", ascending=False).reset_index(drop=True)


def render_sidebar_seller_goals(sellers: pd.DataFrame, key_suffix: str) -> pd.DataFrame:
    st.markdown("### Metas")
    if sellers.empty:
        st.caption("Sem vendedores/canais no período.")
        return seller_goal_table(sellers)

    goals = seller_goal_table(sellers)
    edited = st.data_editor(
        goals[["Vendedor/canal", "Realizado PV", "Meta"]],
        use_container_width=True,
        hide_index=True,
        disabled=["Vendedor/canal", "Realizado PV"],
        column_config={
            "Realizado PV": st.column_config.NumberColumn(format="R$ %.2f"),
            "Meta": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.0, step=1000.0),
        },
        key=f"sidebar_seller_goals_editor_{key_suffix}",
    )
    edited["Meta"] = pd.to_numeric(edited["Meta"], errors="coerce").fillna(0.0)
    st.caption(f"Meta total: {money(float(edited['Meta'].sum()))}")
    return edited


def filter_data(daily: pd.DataFrame, sellers: pd.DataFrame, families: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if daily.empty:
        return daily, sellers, families, seller_goal_table(sellers)

    st.session_state["_demo_month_blocked"] = False
    min_date = daily["data"].min().date()
    max_date = daily["data"].max().date()
    ordered_months = list(daily.sort_values("data")["mes"].drop_duplicates())
    months = ["Todos"] + ordered_months
    seller_names = ["Todos"] + sorted(sellers.loc[sellers["entrada_pv"] > 0, "vendedor"].dropna().unique().tolist())
    family_names = ["Todas"] + sorted(families["familia"].dropna().unique().tolist()) if not families.empty else ["Todas"]
    default_month = months.index("Jun/26") if is_demo_access() and "Jun/26" in months else 0

    with st.sidebar:
        st.header("Filtros")
        selected_month = st.selectbox("Mês", months, index=default_month)
        date_range = st.date_input("Período", value=(min_date, max_date), min_value=min_date, max_value=max_date, format="DD/MM/YYYY")
        selected_seller = st.selectbox("Vendedor / canal de entrada", seller_names)
        selected_family = st.selectbox("Família de produtos", family_names)
        only_business_days = st.checkbox("Somente dias úteis", value=False)

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
    else:
        start, end = pd.to_datetime(min_date), pd.to_datetime(max_date)

    filtered_daily = daily[(daily["data"] >= start) & (daily["data"] <= end)].copy()
    filtered_sellers = sellers[(sellers["data"] >= start) & (sellers["data"] <= end)].copy()
    if families.empty:
        filtered_families = families.copy()
    else:
        start_month = start.to_period("M").to_timestamp()
        end_month = end.to_period("M").to_timestamp()
        filtered_families = families[(families["mes_ref"] >= start_month) & (families["mes_ref"] <= end_month)].copy()

    if selected_month != "Todos":
        month_ref = daily.loc[daily["mes"] == selected_month, "mes_ref"].dropna()
        if is_demo_access() and (
            month_ref.empty or (int(month_ref.iloc[0].year), int(month_ref.iloc[0].month)) not in DEMO_ALLOWED_MONTHS
        ):
            st.session_state["_demo_month_blocked"] = True
            return daily.iloc[0:0].copy(), sellers.iloc[0:0].copy(), families.iloc[0:0].copy(), seller_goal_table(sellers.iloc[0:0].copy())
        filtered_daily = filtered_daily[filtered_daily["mes"] == selected_month]
        filtered_sellers = filtered_sellers[filtered_sellers["mes"] == selected_month]
        filtered_families = filtered_families[filtered_families["mes"] == selected_month]
    elif is_demo_access():
        st.session_state["_demo_month_blocked"] = True
        return daily.iloc[0:0].copy(), sellers.iloc[0:0].copy(), families.iloc[0:0].copy(), seller_goal_table(sellers.iloc[0:0].copy())

    if only_business_days:
        filtered_daily = filtered_daily[filtered_daily["dia_util"] == 1]
        filtered_sellers = filtered_sellers[filtered_sellers["data"].isin(filtered_daily["data"])]

    goals_source = filtered_sellers.copy()
    key_suffix = f"{selected_month}_{start:%Y%m%d}_{end:%Y%m%d}_{int(only_business_days)}"
    with st.sidebar:
        sidebar_seller_goals = render_sidebar_seller_goals(goals_source, key_suffix)

    if selected_seller != "Todos":
        seller_dates = filtered_sellers.loc[
            (filtered_sellers["vendedor"] == selected_seller) & (filtered_sellers["entrada_pv"] > 0),
            "data",
        ]
        filtered_sellers = filtered_sellers[filtered_sellers["vendedor"] == selected_seller]
        filtered_daily = filtered_daily[filtered_daily["data"].isin(seller_dates)]
        sidebar_seller_goals = sidebar_seller_goals[
            sidebar_seller_goals["Vendedor/canal"] == selected_seller
        ].copy()

    if selected_family != "Todas" and not filtered_families.empty:
        filtered_families = filtered_families[filtered_families["familia"] == selected_family]

    return filtered_daily, filtered_sellers, filtered_families, sidebar_seller_goals

def render_overview(daily: pd.DataFrame, sellers: pd.DataFrame, meta_fardin: pd.DataFrame) -> None:
    if daily.empty:
        st.warning("Nenhum dado encontrado para os filtros selecionados.")
        return

    latest_by_month = daily.sort_values("data").groupby("mes_ref", as_index=False).tail(1)
    total_nf = latest_by_month["faturamento_acumulado"].sum()
    total_pv = sellers["entrada_pv"].sum()
    total_pedidos = daily["qtd_pedidos"].sum()
    meta_periodo = latest_by_month["meta_mes"].sum()
    backlog_atual = daily.sort_values("data")["backlog_pedidos"].replace(0, np.nan).dropna()
    backlog_atual_val = float(backlog_atual.iloc[-1]) if not backlog_atual.empty else 0
    ticket = total_pv / total_pedidos if total_pedidos else 0
    atingimento = total_nf / meta_periodo if meta_periodo else 0
    gap = total_nf - meta_periodo

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        metric_card("Faturamento NF acumulado", money(total_nf), pct(atingimento) + " da meta")
    with col2:
        metric_card("Entrada PV", money(total_pv), "Base dos vendedores/canais")
    with col3:
        metric_card("Qtd. pedidos", number(total_pedidos), f"Ticket {money(ticket)}")
    with col4:
        metric_card("Gap vs meta do período", money(gap), money(backlog_atual_val) + " backlog atual")

    col5, col6, col7, col8 = st.columns(4)
    active_days = int((daily["faturamento_realizado"] > 0).sum())
    business_days = int((daily["dia_util"] == 1).sum())
    avg_day = daily.loc[daily["faturamento_realizado"] > 0, "faturamento_realizado"].mean()
    best_day = daily.loc[daily["faturamento_realizado"].idxmax()]
    with col5:
        metric_card("Dias com venda NF", number(active_days), f"{active_days}/{business_days} dias úteis")
    with col6:
        metric_card("Média diária NF", money(avg_day), "Somente dias com venda")
    with col7:
        metric_card("Melhor dia NF", money(best_day["faturamento_realizado"]), best_day["data"].strftime("%d/%m/%Y"))
    with col8:
        realised = meta_fardin["total_real"].sum() if not meta_fardin.empty else 0
        target = meta_fardin["total_meta"].sum() if not meta_fardin.empty else 0
        metric_card("Meta x real Fardin 2026", pct(realised / target if target else 0), f"{money(realised)} realizado")


def render_daily_goals(daily: pd.DataFrame, sellers: pd.DataFrame, families: pd.DataFrame, seller_goals: pd.DataFrame) -> None:
    latest_by_month = daily.sort_values("data").groupby("mes_ref", as_index=False).tail(1)
    realised_total = float(latest_by_month["faturamento_acumulado"].sum()) if not latest_by_month.empty else 0.0
    total_goal = float(seller_goals["Meta"].sum()) if not seller_goals.empty else 0.0

    st.markdown("#### Metas e atingimento")
    col_goal, col_result, col_pct = st.columns(3)
    with col_goal:
        metric_card("Meta total do período", money(total_goal), "Soma das metas por vendedor/canal")
    total_pct = realised_total / total_goal if total_goal else 0.0
    with col_result:
        metric_card("Realizado total", money(realised_total), "Faturamento NF acumulado")
    with col_pct:
        metric_card("Atingimento total", pct(total_pct), money(realised_total - total_goal) + " de gap")

    if seller_goals.empty:
        st.info("Sem vendas por vendedor/canal no período para calcular metas individuais.")
    else:
        edited = seller_goals.copy()
        edited["% ating."] = np.where(edited["Meta"] > 0, edited["Realizado PV"] / edited["Meta"], 0)
        edited["Gap"] = edited["Realizado PV"] - edited["Meta"]
        render_table(
            edited,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Realizado PV": st.column_config.NumberColumn(format="R$ %.2f"),
                "Meta": st.column_config.NumberColumn(format="R$ %.2f"),
                "% ating.": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=1),
                "Gap": st.column_config.NumberColumn(format="R$ %.2f"),
            },
        )

    if families.empty:
        st.info("Sem dados de família de produtos para os filtros selecionados.")
        return

    st.markdown("#### Meta x realizado por família")
    family_summary = (
        families.groupby("familia", as_index=False)
        .agg(meta_familia=("meta_familia", "sum"), realizado_familia=("realizado_familia", "sum"))
        .sort_values("realizado_familia", ascending=False)
    )
    family_summary["atingimento_familia"] = np.where(
        family_summary["meta_familia"] > 0,
        family_summary["realizado_familia"] / family_summary["meta_familia"],
        0,
    )
    family_summary["gap"] = family_summary["realizado_familia"] - family_summary["meta_familia"]

    fig_family = px.bar(
        family_summary,
        x="familia",
        y=["realizado_familia", "meta_familia"],
        barmode="group",
        title="Famílias de produtos - meta x realizado",
        labels={"value": "R$", "familia": "Família", "variable": "Indicador"},
    )
    plot_chart(fig_family, use_container_width=True)
    render_table(
        family_summary.rename(
            columns={
                "familia": "Família",
                "meta_familia": "Meta",
                "realizado_familia": "Realizado",
                "atingimento_familia": "% ating.",
                "gap": "Gap",
            }
        ),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Meta": st.column_config.NumberColumn(format="R$ %.2f"),
            "Realizado": st.column_config.NumberColumn(format="R$ %.2f"),
            "% ating.": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=1),
            "Gap": st.column_config.NumberColumn(format="R$ %.2f"),
        },
    )

def render_daily_tab(daily: pd.DataFrame, sellers: pd.DataFrame, metas: pd.DataFrame, families: pd.DataFrame, seller_goals: pd.DataFrame) -> None:
    st.subheader("Resumo diário de vendas")
    st.markdown('<div class="section-note">Informações vindas das abas mensais do arquivo de Resumo Diário.</div>', unsafe_allow_html=True)

    if daily.empty:
        st.info("Sem dados no período.")
        return

    render_daily_goals(daily, sellers, families, seller_goals)

    by_day = daily.sort_values("data")
    fig = px.bar(
        by_day,
        x="data",
        y="faturamento_realizado",
        color="mes",
        title="Faturamento realizado por dia",
        labels={"data": "Data", "faturamento_realizado": "Faturamento NF"},
    )
    fig.update_xaxes(tickformat="%d/%m/%Y")
    fig.add_scatter(
        x=by_day["data"],
        y=by_day["resultado_projetado"],
        mode="lines",
        name="Resultado projetado",
        line=dict(color="#dc2626", width=2),
    )
    plot_chart(fig, use_container_width=True)

    col1, col2 = st.columns([1.2, 1])
    with col1:
        month_last = daily.sort_values("data").groupby("mes_ref", as_index=False).tail(1)
        month_last = month_last.assign(
            atingimento_meta=lambda d: np.where(d["meta_mes"] > 0, d["faturamento_acumulado"] / d["meta_mes"], 0),
            falta_meta=lambda d: d["meta_mes"] - d["faturamento_acumulado"],
        )
        fig_month = px.bar(
            month_last,
            x="mes",
            y=["faturamento_acumulado", "meta_mes"],
            barmode="group",
            title="Meta x realizado por mês",
            labels={"value": "R$", "variable": "Indicador"},
        )
        plot_chart(fig_month, use_container_width=True)
    with col2:
        render_table(
            month_last[
                ["mes", "meta_mes", "faturamento_acumulado", "atingimento_meta", "falta_meta", "dias_uteis_mes"]
            ].rename(
                columns={
                    "mes": "Mês",
                    "meta_mes": "Meta mês",
                    "faturamento_acumulado": "Realizado NF",
                    "atingimento_meta": "% ating.",
                    "falta_meta": "Falta meta",
                    "dias_uteis_mes": "Dias úteis",
                }
            ),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Meta mês": st.column_config.NumberColumn(format="R$ %.2f"),
                "Realizado NF": st.column_config.NumberColumn(format="R$ %.2f"),
                "% ating.": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=1),
                "Falta meta": st.column_config.NumberColumn(format="R$ %.2f"),
            },
        )

    daily_view = daily[
        [
            "data",
            "dia_semana",
            "dia_util",
            "resultado_projetado",
            "faturamento_realizado",
            "faturamento_acumulado",
            "atingimento_dia",
            "qtd_pedidos",
            "entrada_pv",
            "ticket_medio",
            "backlog_pedidos",
            "ajuste",
        ]
    ].rename(
        columns={
            "data": "Data",
            "dia_semana": "Dia semana",
            "dia_util": "Dia útil",
            "resultado_projetado": "Resultado projetado",
            "faturamento_realizado": "Faturamento realizado",
            "faturamento_acumulado": "Faturamento acumulado",
            "atingimento_dia": "Projetado x realizado",
            "qtd_pedidos": "Qtd pedidos",
            "entrada_pv": "Entrada PV",
            "ticket_medio": "Ticket médio",
            "backlog_pedidos": "Backlog pedidos",
            "ajuste": "Ajuste",
        }
    )
    daily_view["Data"] = br_date_series(daily_view["Data"])
    render_table(daily_view, use_container_width=True, hide_index=True)


def render_seller_tab(sellers: pd.DataFrame) -> None:
    st.subheader("Entrada por vendedor / canal")
    if sellers.empty:
        st.info("Sem dados de vendedores no período.")
        return

    seller_rank = (
        sellers.groupby("vendedor", as_index=False)
        .agg(entrada_pv=("entrada_pv", "sum"), dias_com_entrada=("entrada_pv", lambda s: int((s > 0).sum())))
        .sort_values("entrada_pv", ascending=False)
    )
    seller_rank["participacao"] = seller_rank["entrada_pv"] / seller_rank["entrada_pv"].sum()

    col1, col2 = st.columns([1.35, 1])
    with col1:
        fig = px.bar(
            seller_rank,
            x="entrada_pv",
            y="vendedor",
            orientation="h",
            title="Ranking de entrada PV por vendedor/canal",
            labels={"entrada_pv": "Entrada PV", "vendedor": ""},
        )
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
        plot_chart(fig, use_container_width=True)
    with col2:
        fig_pie = px.pie(
            seller_rank,
            names="vendedor",
            values="entrada_pv",
            title="Participação na entrada PV",
            hole=0.45,
        )
        plot_chart(fig_pie, use_container_width=True)

    by_month = sellers.groupby(["mes", "vendedor"], as_index=False)["entrada_pv"].sum()
    fig_month = px.bar(
        by_month,
        x="mes",
        y="entrada_pv",
        color="vendedor",
        title="Entrada PV por mês e vendedor/canal",
        labels={"entrada_pv": "Entrada PV", "mes": "Mês"},
    )
    plot_chart(fig_month, use_container_width=True)

    render_table(
        seller_rank.rename(
            columns={
                "vendedor": "Vendedor/canal",
                "entrada_pv": "Entrada PV",
                "dias_com_entrada": "Dias com entrada",
                "participacao": "Participação",
            }
        ),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Entrada PV": st.column_config.NumberColumn(format="R$ %.2f"),
            "Participação": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=1),
        },
    )


def render_goals_tab(meta_global: pd.DataFrame, meta_fardin: pd.DataFrame, evo_global: pd.DataFrame, evo_fardin: pd.DataFrame) -> None:
    st.subheader("Metas, realizado e evolução")

    col1, col2 = st.columns(2)
    with col1:
        fig_global = px.bar(
            meta_global,
            x="mes",
            y=["meta", "total"],
            barmode="group",
            title="Global - meta x realizado",
            labels={"value": "R$", "variable": "Indicador"},
        )
        plot_chart(fig_global, use_container_width=True)
    with col2:
        fig_fardin = px.bar(
            meta_fardin,
            x="mes",
            y=["total_meta", "total_real"],
            barmode="group",
            title="Fardin - meta x realizado",
            labels={"value": "R$", "variable": "Indicador"},
        )
        plot_chart(fig_fardin, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        fig_cat = px.bar(
            meta_global,
            x="mes",
            y=["fardin", "privaty_label", "terceiros"],
            title="Global por categoria de vendas",
            labels={"value": "R$", "variable": "Categoria"},
        )
        plot_chart(fig_cat, use_container_width=True)
    with col4:
        fig_rep = px.bar(
            meta_fardin,
            x="mes",
            y=["representantes_real", "doces_fardin_real", "wilson_real"],
            title="Fardin por grupo",
            labels={"value": "R$", "variable": "Grupo"},
        )
        plot_chart(fig_rep, use_container_width=True)

    evo_global_long = evo_global.melt(id_vars=["mes"], value_vars=["base_vendas", "real_vendas"], var_name="periodo", value_name="vendas")
    evo_fardin_long = evo_fardin.melt(id_vars=["mes"], value_vars=["base_vendas", "real_vendas"], var_name="periodo", value_name="vendas")
    col5, col6 = st.columns(2)
    with col5:
        plot_chart(px.line(evo_global_long, x="mes", y="vendas", color="periodo", markers=True, title="Evolução GLOBAL 2025 x 2026"), use_container_width=True)
    with col6:
        plot_chart(px.line(evo_fardin_long, x="mes", y="vendas", color="periodo", markers=True, title="Evolução FARDIN 2025 x 2026"), use_container_width=True)

    st.markdown("#### Tabela - Meta x Real Global")
    render_table(meta_global, use_container_width=True, hide_index=True)
    st.markdown("#### Tabela - Meta x Real Fardin")
    render_table(meta_fardin, use_container_width=True, hide_index=True)


def render_history_tab(hist_pf: pd.DataFrame, hist_geral: pd.DataFrame, hist_reps: pd.DataFrame, budget: pd.DataFrame) -> None:
    st.subheader("Histórico, famílias, representantes e orçamento")

    if not hist_pf.empty:
        real_family = hist_pf[hist_pf["indicador"].eq("Real")].copy()
        real_family = real_family[real_family["valor"] > 0]
        fig_family = px.bar(
            real_family,
            x="mes",
            y="valor",
            color="familia",
            title="Faturamento por família",
            labels={"valor": "R$", "mes": "Mês"},
        )
        plot_chart(fig_family, use_container_width=True)

        ating = hist_pf[hist_pf["indicador"].eq("% de Atingimento")].copy()
        if not ating.empty:
            fig_ating = px.line(
                ating,
                x="mes",
                y="valor",
                color="familia",
                markers=True,
                title="% de atingimento por família",
                labels={"valor": "% atingimento", "mes": "Mês"},
            )
            plot_chart(fig_ating, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        if not hist_geral.empty:
            fat = hist_geral[hist_geral["indicador"].eq("Faturamento")]
            plot_chart(px.bar(fat, x="mes", y="valor", title="Histórico geral - faturamento"), use_container_width=True)
    with col2:
        if not hist_reps.empty:
            nf = hist_reps[hist_reps["indicador"].eq("Faturamento (NF)")].copy()
            nf = nf[nf["valor"] > 0]
            plot_chart(px.bar(nf, x="mes", y="valor", color="representante", title="Faturamento NF por representante"), use_container_width=True)

    if not budget.empty:
        budget_long = budget.melt(
            id_vars=["mes"],
            value_vars=["viagens_planejado", "viagens_realizado", "pdv_planejado", "pdv_realizado", "eventos_planejado", "eventos_realizado"],
            var_name="indicador",
            value_name="valor",
        )
        plot_chart(
            px.bar(budget_long, x="mes", y="valor", color="indicador", barmode="group", title="Orçamento comercial - planejado x realizado"),
            use_container_width=True,
        )

    st.markdown("#### Histórico por família")
    render_table(hist_pf, use_container_width=True, hide_index=True)
    st.markdown("#### Histórico geral")
    render_table(hist_geral, use_container_width=True, hide_index=True)
    st.markdown("#### Histórico por representante")
    render_table(hist_reps, use_container_width=True, hide_index=True)
    st.markdown("#### Orçamento comercial")
    render_table(budget, use_container_width=True, hide_index=True)



def render_impact_per_capita_tab(
    agenda_points: pd.DataFrame,
    agenda_calendar: pd.DataFrame,
    per_capita_detail: pd.DataFrame,
    per_capita_summary: pd.DataFrame,
    graficos_raw: pd.DataFrame,
) -> None:
    st.subheader("Agenda de impacto e faturamento per capita")

    st.markdown("#### Agenda de Impacto")
    col1, col2 = st.columns([0.9, 1.5])
    with col1:
        if agenda_points.empty:
            st.info("Nenhum ponto de impacto textual encontrado na agenda.")
        else:
            render_table(
                agenda_points.rename(columns={"ponto_de_impacto": "Ponto importante"}),
                use_container_width=True,
                hide_index=True,
            )
    with col2:
        st.caption("Calendario comercial preservado para conferencia visual da agenda original.")
        render_table(agenda_calendar, use_container_width=True, hide_index=True)

    st.markdown("#### Faturamento per Capita")
    if per_capita_detail.empty:
        st.info("Sem dados de faturamento per capita para exibir.")
    else:
        col3, col4, col5 = st.columns(3)
        current = per_capita_summary[per_capita_summary["periodo"].eq("2024")]
        base = per_capita_summary[per_capita_summary["periodo"].eq("2023")]
        current_pc = float(current["per_capita"].iloc[0]) if not current.empty else 0
        base_pc = float(base["per_capita"].iloc[0]) if not base.empty else 0
        current_sales = float(current["vendas"].iloc[0]) if not current.empty else 0
        current_people = float(current["funcionarios"].iloc[0]) if not current.empty else 0
        col3.metric("Per capita medio atual", money(current_pc), pct((current_pc / base_pc - 1) if base_pc else 0))
        col4.metric("Venda media mensal atual", money(current_sales))
        col5.metric("Funcionarios medio atual", number(current_people, 1))

        pc_plot = per_capita_detail[per_capita_detail["data"].notna()].copy()
        fig_pc = px.line(
            pc_plot,
            x="data",
            y="per_capita",
            color="periodo",
            markers=True,
            title="R$ per capita por mes",
            labels={"data": "Mes", "per_capita": "R$ per capita"},
        )
        fig_pc.update_xaxes(tickformat="%m/%Y")
        plot_chart(fig_pc, use_container_width=True)

        col6, col7 = st.columns(2)
        with col6:
            fig_sales = px.bar(
                pc_plot,
                x="data",
                y="vendas",
                color="periodo",
                title="Faturamento usado no per capita",
                labels={"data": "Mes", "vendas": "R$ vendas"},
            )
            fig_sales.update_xaxes(tickformat="%m/%Y")
            plot_chart(fig_sales, use_container_width=True)
        with col7:
            fig_gain = px.bar(
                pc_plot[pc_plot["ganho"] != 0],
                x="data",
                y="ganho",
                color="periodo",
                title="Ganho / perda vs base",
                labels={"data": "Mes", "ganho": "R$"},
            )
            fig_gain.update_xaxes(tickformat="%m/%Y")
            plot_chart(fig_gain, use_container_width=True)

        per_capita_view = per_capita_detail.rename(
            columns={
                "periodo": "Período",
                "data": "Data",
                "mes": "Mês",
                "vendas": "Vendas",
                "funcionarios": "Funcionários",
                "per_capita": "R$ per capita",
                "ganho": "Ganho / perda",
                "acumulado": "Acumulado",
                "media_mensal": "Média mensal",
                "anualizado": "Anualizado",
            }
        )
        per_capita_view["Data"] = br_date_series(per_capita_view["Data"])
        render_table(per_capita_view, use_container_width=True, hide_index=True)

    st.markdown("#### Aba Graficos")
    st.caption("A aba 'Graficos' do Excel nao contem serie de dados; ela traz apenas o titulo do painel. Por isso fica aqui como conferencia, sem grafico derivado.")
    render_table(graficos_raw, use_container_width=True, hide_index=True)
def render_raw_tabs(resumo_source: WorkbookSource, avaliacoes_source: WorkbookSource) -> None:
    st.subheader("Conferência das planilhas originais")
    st.caption("Visualização simplificada das abas originais para auditoria das informações que alimentam o dashboard.")
    file_choice = st.selectbox("Arquivo", ["Resumo Diário", "Avaliações de Desempenho"])
    path = resumo_source if file_choice == "Resumo Diário" else avaliacoes_source
    xls = excel_file(path)
    sheet = st.selectbox("Aba", xls.sheet_names)
    raw = read_excel_workbook(path, sheet_name=sheet, header=None)
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    render_table(raw, use_container_width=True, hide_index=True)


def render_footer() -> None:
    st.markdown("<div class='app-footer'>", unsafe_allow_html=True)
    col_text, col_logo = st.columns([0.78, 0.22], vertical_alignment="center")
    with col_text:
        st.markdown(
            """
            <div class="footer-copy">
              <strong>Desenvolvido por Fabrício Zamprogno</strong>
              <span>em parceria com Alfa Soluções Consultoria</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_logo:
        if ALFA_LOGO_PATH.exists():
            st.image(str(ALFA_LOGO_PATH), width=118)
    st.markdown("</div>", unsafe_allow_html=True)

def render_development_notice(title: str = "Modulo em desenvolvimento") -> None:
    st.markdown(
        f"""
        <div style="
            margin-top: 18px;
            padding: 22px;
            border: 1px solid #334155;
            border-radius: 8px;
            background: rgba(17, 24, 39, .92);
        ">
          <h3 style="margin:0 0 8px 0;color:#f8fafc">{title}</h3>
          <p style="margin:0;color:#cbd5e1">
            Esta area faz parte da proxima etapa de implantacao do dashboard comercial.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

def main() -> None:
    require_login()
    theme = select_theme()
    st.session_state["_theme"] = theme
    pio.templates.default = 'plotly_dark' if theme == 'dark' else 'plotly_white'
    if theme == 'light':
        pio.templates['plotly_white'].layout.update(
            paper_bgcolor='#ffffff',
            plot_bgcolor='#ffffff',
            font_color='#111827',
        )
    style_app(theme)
    resumo_source, avaliacoes_source, resumo_name, avaliacoes_name = select_data_sources()
    if LOGO_PATH.exists():
        logo_col, title_col = st.columns([0.11, 0.89], vertical_alignment="center")
        with logo_col:
            st.image(str(LOGO_PATH), width=104)
        with title_col:
            st.markdown(
                """
                <div class="hero">
                  <h1>Dashboard Comercial - Doces Fardin 2026</h1>
                  <p>Resumo diario, metas, desempenho comercial, historico por familia, representantes e orcamento.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
    else:
        st.markdown(
            """
            <div class="hero">
              <h1>Dashboard Comercial - Doces Fardin 2026</h1>
              <p>Resumo diario, metas, desempenho comercial, historico por familia, representantes e orcamento.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if resumo_source is None or avaliacoes_source is None:
        st.warning("Envie os dois arquivos Excel para carregar o dashboard.")
        st.markdown(
            """
            <div class="file-required-panel">
              <strong>Arquivos obrigatórios</strong>
              <span>1. Resumo Diário de Vendas</span>
              <span>2. Avaliações de Desempenho Comercial</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.stop()

    missing = []
    if isinstance(resumo_source, Path) and not resumo_source.exists():
        missing.append(str(resumo_source))
    if isinstance(avaliacoes_source, Path) and not avaliacoes_source.exists():
        missing.append(str(avaliacoes_source))
    if missing:
        st.error("Arquivo(s) nao encontrado(s): " + ", ".join(missing))
        st.info("Selecione os dois arquivos Excel na barra lateral em Arquivos de dados.")
        st.stop()
    try:
        daily, sellers, metas = load_monthly_summary(resumo_source)
        meta_global = load_meta_global(avaliacoes_source)
        meta_fardin = load_meta_fardin(avaliacoes_source)
        evo_global = load_evolution(avaliacoes_source, "Evolução de Vendas Mês GLOBAL")
        evo_fardin = load_evolution(avaliacoes_source, "Evolução de Vendas Mês FARDIN")
        hist_pf = load_history_pf(avaliacoes_source)
        hist_geral, hist_reps = load_history_pr(avaliacoes_source)
        budget = load_budget(avaliacoes_source)
        agenda_points, agenda_calendar = load_agenda_impact(avaliacoes_source)
        per_capita_detail, per_capita_summary = load_per_capita(avaliacoes_source)
        graficos_raw = load_graficos_raw(avaliacoes_source)
        family_performance = load_family_performance(avaliacoes_source)
    except Exception as exc:
        st.error("Nao foi possivel ler os arquivos selecionados.")
        st.caption(f"Resumo: {resumo_name} | Avaliacoes: {avaliacoes_name}")
        st.exception(exc)
        st.stop()

    filtered_daily, filtered_sellers, filtered_families, seller_goals = filter_data(daily, sellers, family_performance)

    if is_demo_access() and st.session_state.get("_demo_month_blocked"):
        render_development_notice("Mes em desenvolvimento")
        st.stop()

    render_overview(filtered_daily, filtered_sellers, meta_fardin)

    tab_daily, tab_sellers, tab_goals, tab_history, tab_impact, tab_raw = st.tabs(
        [
            "Resumo diário",
            "Vendedores e canais",
            "Metas e evolução",
            "Histórico e orçamento",
            "Agenda e per capita",
            "Conferência Excel",
        ]
    )
    with tab_daily:
        render_daily_tab(filtered_daily, filtered_sellers, metas, filtered_families, seller_goals)
    with tab_sellers:
        render_seller_tab(filtered_sellers)
    with tab_goals:
        if is_demo_access():
            render_development_notice()
        else:
            render_goals_tab(meta_global, meta_fardin, evo_global, evo_fardin)
    with tab_history:
        if is_demo_access():
            render_development_notice()
        else:
            render_history_tab(hist_pf, hist_geral, hist_reps, budget)
    with tab_impact:
        if is_demo_access():
            render_development_notice()
        else:
            render_impact_per_capita_tab(
                agenda_points,
                agenda_calendar,
                per_capita_detail,
                per_capita_summary,
                graficos_raw,
            )
    with tab_raw:
        if is_demo_access():
            render_development_notice()
        else:
            render_raw_tabs(resumo_source, avaliacoes_source)

    render_footer()


if __name__ == "__main__":
    main()
