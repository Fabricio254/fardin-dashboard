from __future__ import annotations

from base64 import b64encode
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
PEDIDO_XLS = BASE_DIR / "Pedido.XLS"
VENDA_XLS = BASE_DIR / "Venda.XLS"
LOGO_PATH = BASE_DIR / "logo fardin.jpg"
FAVICON_PATH = BASE_DIR / "favicon-fardin.png"

WorkbookSource = str | Path | bytes


st.set_page_config(
    page_title="Dashboard Comercial - Fardin",
    page_icon=str(FAVICON_PATH),
    layout="wide",
)


def style_app() -> None:
    st.markdown(
        """
        <style>
        #MainMenu,
        footer,
        header,
        [data-testid="stToolbar"],
        [data-testid="stHeader"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        [data-testid="stDeployButton"],
        [data-testid="stAppDeployButton"],
        [data-testid="manage-app-button"],
        .stDeployButton,
        .viewerBadge_container__1QSob,
        .viewerBadge_link__1S137,
        .viewerBadge_text__1JaDK {
            display: none !important;
            visibility: hidden !important;
            height: 0 !important;
        }
        .stApp {
            margin-top: 0 !important;
        }
        .block-container {
            padding-top: 2.8rem;
            max-width: 1220px;
        }
        .fardin-header {
            margin-bottom: 1.1rem;
        }
        .logo-frame {
            width: 96px;
            height: 96px;
            padding: 6px;
            border-radius: 8px;
            background: rgba(255, 255, 255, .06);
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: visible;
        }
        .logo-frame img {
            width: 84px;
            height: 84px;
            object-fit: contain;
            display: block;
        }
        .fardin-header h1 {
            margin: 0;
            font-size: 2.35rem;
            line-height: 1.05;
            letter-spacing: 0;
        }
        .kpi-card {
            min-height: 116px;
            padding: 15px 15px 13px;
            border: 1px solid rgba(148, 163, 184, .22);
            border-radius: 8px;
            background:
                linear-gradient(180deg, rgba(31, 41, 55, .95), rgba(15, 23, 42, .96));
            box-shadow: 0 16px 34px rgba(0, 0, 0, .24);
            position: relative;
            overflow: hidden;
        }
        .kpi-card::before {
            content: "";
            position: absolute;
            inset: 0 0 auto 0;
            height: 3px;
            background: var(--accent);
        }
        .kpi-label {
            color: #a7b2c3;
            font-size: .68rem;
            font-weight: 700;
            letter-spacing: .04em;
            text-transform: uppercase;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .kpi-value {
            margin-top: 11px;
            color: #f8fafc;
            font-size: clamp(1.02rem, 1.38vw, 1.34rem);
            font-weight: 800;
            line-height: 1.16;
            letter-spacing: 0;
            white-space: nowrap;
        }
        .kpi-note {
            margin-top: 8px;
            color: #8fa0b7;
            font-size: .68rem;
        }
        .kpi-card.primary {
            background:
                linear-gradient(180deg, rgba(127, 29, 29, .42), rgba(15, 23, 42, .98)),
                linear-gradient(180deg, rgba(31, 41, 55, .96), rgba(15, 23, 42, .96));
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid rgba(148, 163, 184, .18);
            border-radius: 8px;
            overflow: hidden;
        }
        h2, h3 {
            letter-spacing: 0;
        }
        section[data-testid="stSidebar"] h2 {
            font-size: 1.05rem;
            margin-bottom: .8rem;
        }
        .upload-heading {
            margin: 1rem 0 .45rem;
            color: #f8fafc;
            font-size: .78rem;
            font-weight: 800;
            letter-spacing: .05em;
            text-transform: uppercase;
        }
        .upload-status {
            margin: .45rem 0 1.05rem;
            padding: 10px 12px;
            border: 1px solid rgba(148, 163, 184, .22);
            border-radius: 8px;
            background: rgba(15, 23, 42, .72);
        }
        .upload-status span {
            display: block;
            color: #94a3b8;
            font-size: .68rem;
            font-weight: 700;
            letter-spacing: .04em;
            text-transform: uppercase;
        }
        .upload-status strong {
            display: block;
            margin-top: 3px;
            color: #f8fafc;
            font-size: .84rem;
            line-height: 1.25;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .upload-status.ready {
            border-color: rgba(34, 197, 94, .38);
            background: linear-gradient(180deg, rgba(22, 101, 52, .20), rgba(15, 23, 42, .78));
        }
        section[data-testid="stSidebar"] div[data-testid="stFileUploader"] {
            margin-bottom: .35rem;
        }
        section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section {
            min-height: 52px;
            padding: 10px 12px;
            border: 1px solid rgba(148, 163, 184, .24);
            border-radius: 8px;
            background: linear-gradient(180deg, rgba(15, 23, 42, .78), rgba(2, 6, 23, .55));
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, .04);
        }
        section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section:hover {
            border-color: rgba(248, 250, 252, .42);
            background: rgba(15, 23, 42, .82);
        }
        section[data-testid="stSidebar"] div[data-testid="stFileUploader"] section > div:first-child {
            display: none;
        }
        section[data-testid="stSidebar"] div[data-testid="stFileUploader"] small {
            display: none;
        }
        section[data-testid="stSidebar"] div[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] {
            display: none;
        }
        section[data-testid="stSidebar"] div[data-testid="stFileUploader"] button {
            min-height: 32px;
            border-radius: 7px;
            border: 1px solid rgba(248, 250, 252, .18);
            background: rgba(248, 250, 252, .08);
            color: #f8fafc;
            font-size: .78rem;
            font-weight: 700;
        }
        .filter-heading {
            margin: 1rem 0 .45rem;
            color: #f8fafc;
            font-size: .78rem;
            font-weight: 800;
            letter-spacing: .05em;
            text-transform: uppercase;
        }
        .filter-summary {
            margin: .45rem 0 .7rem;
            padding: 9px 11px;
            border: 1px solid rgba(56, 189, 248, .22);
            border-radius: 8px;
            background: rgba(15, 23, 42, .72);
            color: #cbd5e1;
            font-size: .76rem;
        }
        section[data-testid="stSidebar"] div[data-testid="stPills"] {
            padding: 10px;
            border: 1px solid rgba(148, 163, 184, .24);
            border-radius: 8px;
            background: linear-gradient(180deg, rgba(15, 23, 42, .86), rgba(2, 6, 23, .68));
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, .04);
        }
        section[data-testid="stSidebar"] div[data-testid="stPills"] button {
            max-width: 100%;
            margin: 0 4px 7px 0;
            border: 1px solid rgba(148, 163, 184, .22);
            border-radius: 7px;
            background: rgba(30, 41, 59, .72);
            color: #dbe7f3;
            font-size: .68rem;
            font-weight: 750;
            letter-spacing: 0;
        }
        section[data-testid="stSidebar"] div[data-testid="stPills"] button:hover {
            border-color: rgba(56, 189, 248, .45);
            background: rgba(51, 65, 85, .92);
            color: #f8fafc;
        }
        section[data-testid="stSidebar"] div[data-testid="stPills"] button[aria-pressed="true"] {
            border-color: rgba(56, 189, 248, .58);
            background: linear-gradient(180deg, rgba(14, 116, 144, .55), rgba(15, 23, 42, .88));
            color: #f8fafc;
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, .08);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def workbook_source(source: WorkbookSource) -> BytesIO | str | Path:
    return BytesIO(source) if isinstance(source, bytes) else source


def money(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        value = 0
    return "R$ " + f"{float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(value: float | int | None) -> str:
    if value is None or pd.isna(value):
        value = 0
    return f"{float(value) * 100:,.1f}%".replace(",", "X").replace(".", ",")


def kpi_card(label: str, value: str, note: str, accent: str, primary: bool = False) -> None:
    class_name = "kpi-card primary" if primary else "kpi-card"
    st.markdown(
        f"""
        <div class="{class_name}" style="--accent:{accent}">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def logo_html(path: Path) -> str:
    encoded = b64encode(path.read_bytes()).decode("ascii")
    return f'<div class="logo-frame"><img src="data:image/jpeg;base64,{encoded}" alt="Fardin"></div>'


def parse_money_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    text = series.fillna("0").astype(str).str.strip()
    text = text.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(text, errors="coerce").fillna(0.0)


def parse_date_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, dayfirst=True, errors="coerce")


def read_xls_table(source: WorkbookSource) -> pd.DataFrame:
    try:
        df = pd.read_excel(workbook_source(source), sheet_name=0, header=1, engine="xlrd")
    except ImportError as exc:
        raise RuntimeError(
            "Para ler arquivos .XLS, instale as dependencias com: pip install -r requirements.txt"
        ) from exc
    except Exception as exc:
        raise RuntimeError(f"Nao foi possivel ler o arquivo Excel: {exc}") from exc

    df = df.dropna(how="all")
    df.columns = [str(col).strip() if not str(col).startswith("Unnamed") else "" for col in df.columns]
    return df


@st.cache_data(show_spinner=False)
def load_pedidos(source: WorkbookSource) -> pd.DataFrame:
    df = read_xls_table(source).copy()
    required = ["Emissao", "Nome do Cliente", "Pedido", "Valor do Pedido", "Valor A Faturar", "Situacao Atendimento", "Vendedor"]
    rename = {
        "Emissão": "Emissao",
        "Situação Atendimento": "Situacao Atendimento",
    }
    df = df.rename(columns=rename)
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError("Pedido.XLS sem coluna(s): " + ", ".join(missing))

    df["Data"] = parse_date_series(df["Emissao"])
    df = df[df["Data"].notna()].copy()
    df["Valor do Pedido"] = parse_money_series(df["Valor do Pedido"])
    df["Valor A Faturar"] = parse_money_series(df["Valor A Faturar"])
    df["Vendedor"] = df["Vendedor"].fillna("").astype(str).str.strip()
    df["Cliente"] = df["Nome do Cliente"].fillna("").astype(str).str.strip()
    df["Situacao Atendimento"] = df["Situacao Atendimento"].fillna("").astype(str).str.strip()
    return df


@st.cache_data(show_spinner=False)
def load_vendas(source: WorkbookSource) -> pd.DataFrame:
    df = read_xls_table(source).copy()
    required = [
        "Nota Fiscal",
        "Nome da Pessoa",
        "Data Emissão",
        "Data Movimento",
        "Valor Total Líquido",
        "Cancelada?",
        "Nome do Vendedor Representante",
    ]
    rename = {
        "Data Emissão": "Data Emissao",
        "Valor Total Líquido": "Valor Total Liquido",
        "Nome da Pessoa": "Cliente",
        "Nome do Vendedor Representante": "Vendedor",
    }
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError("Venda.XLS sem coluna(s): " + ", ".join(missing))

    df = df.rename(columns=rename)
    df["Data"] = parse_date_series(df["Data Emissao"])
    df = df[df["Data"].notna()].copy()
    df["Valor Total Liquido"] = parse_money_series(df["Valor Total Liquido"])
    df["Vendedor"] = df["Vendedor"].fillna("").astype(str).str.strip()
    df["Cliente"] = df["Cliente"].fillna("").astype(str).str.strip()
    df["Cancelada?"] = df["Cancelada?"].fillna("").astype(str).str.strip()
    df["Venda Valida"] = df["Cancelada?"].str.casefold() != "verdadeiro"
    return df


def source_selector(label: str, default_path: Path, key: str) -> tuple[WorkbookSource | None, str]:
    st.markdown(f'<div class="upload-heading">{label}</div>', unsafe_allow_html=True)
    upload = st.file_uploader(
        f"Arquivo {label}",
        type=["xls", "xlsx"],
        key=key,
        label_visibility="collapsed",
    )

    if upload is not None:
        source: WorkbookSource | None = upload.getvalue()
        name = upload.name
        status = "Selecionado"
        status_class = "ready"
    elif default_path.exists():
        source = default_path
        name = default_path.name
        status = "Padrao local"
        status_class = "ready"
    else:
        source = None
        name = "Aguardando arquivo"
        status = "Pendente"
        status_class = ""

    st.markdown(
        f"""
        <div class="upload-status {status_class}">
          <span>{status}</span>
          <strong>{name}</strong>
        </div>
        """,
        unsafe_allow_html=True,
    )
    return source, name


def build_meta_editor(vendedores: list[str], total_vendido: float) -> tuple[float, pd.DataFrame]:
    st.sidebar.markdown("### Metas")
    meta_total = st.sidebar.number_input(
        "Meta total do periodo",
        min_value=0.0,
        value=float(round(total_vendido, 2)) if total_vendido else 0.0,
        step=1000.0,
        format="%.2f",
    )

    base = pd.DataFrame({"Vendedor": vendedores, "Meta": [0.0] * len(vendedores)})
    if "metas_vendedores" not in st.session_state:
        st.session_state["metas_vendedores"] = base
    else:
        old = st.session_state["metas_vendedores"]
        known = dict(zip(old.get("Vendedor", []), old.get("Meta", [])))
        st.session_state["metas_vendedores"] = pd.DataFrame(
            {"Vendedor": vendedores, "Meta": [float(known.get(v, 0.0) or 0.0) for v in vendedores]}
        )

    metas = st.sidebar.data_editor(
        st.session_state["metas_vendedores"],
        key="metas_editor",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Vendedor": st.column_config.TextColumn("Vendedor", disabled=True),
            "Meta": st.column_config.NumberColumn("Meta", min_value=0.0, step=500.0, format="R$ %.2f"),
        },
    )
    return meta_total, metas


def apply_filters(pedidos: pd.DataFrame, vendas: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    all_dates = pd.concat([pedidos["Data"], vendas["Data"]]).dropna()
    min_date = all_dates.min().date()
    max_date = all_dates.max().date()

    st.sidebar.markdown("### Filtros")
    date_range = st.sidebar.date_input("Periodo de venda", value=(min_date, max_date), min_value=min_date, max_value=max_date, format="DD/MM/YYYY")
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
    else:
        start = end = date_range

    vendedores = sorted(set(pedidos["Vendedor"]).union(vendas["Vendedor"]) - {""})
    st.sidebar.markdown('<div class="filter-heading">Vendedores</div>', unsafe_allow_html=True)
    selected = st.sidebar.pills(
        "Vendedores",
        vendedores,
        selection_mode="multi",
        default=vendedores,
        label_visibility="collapsed",
        width="stretch",
    )
    selected = selected or []
    st.sidebar.markdown(
        f'<div class="filter-summary">{len(selected)} de {len(vendedores)} vendedores selecionados</div>',
        unsafe_allow_html=True,
    )

    start_ts = pd.Timestamp(start)
    end_ts = pd.Timestamp(end)
    pedidos_f = pedidos[pedidos["Data"].between(start_ts, end_ts)].copy()
    vendas_f = vendas[vendas["Data"].between(start_ts, end_ts)].copy()
    pedidos_f = pedidos_f[pedidos_f["Vendedor"].isin(selected)]
    vendas_f = vendas_f[vendas_f["Vendedor"].isin(selected)]
    return pedidos_f, vendas_f


def render_kpis(pedidos: pd.DataFrame, vendas: pd.DataFrame, meta_total: float) -> None:
    vendas_validas = vendas[vendas["Venda Valida"]]
    total_vendido = float(vendas_validas["Valor Total Liquido"].sum())
    total_pedido = float(pedidos["Valor do Pedido"].sum())
    total_a_faturar = float(pedidos["Valor A Faturar"].sum())
    atingimento = total_vendido / meta_total if meta_total else 0.0
    pedidos_qtd = int(pedidos["Pedido"].nunique()) if not pedidos.empty else 0
    notas_qtd = int(vendas_validas["Nota Fiscal"].nunique()) if not vendas_validas.empty else 0

    cols = st.columns(5)
    with cols[0]:
        kpi_card("Vendas validas", money(total_vendido), f"{notas_qtd} notas no periodo", "#22c55e", primary=True)
    with cols[1]:
        kpi_card("Meta total", money(meta_total), "Informada na barra lateral", "#f59e0b")
    with cols[2]:
        kpi_card("Atingimento", pct(atingimento), "Vendas validas / meta", "#38bdf8")
    with cols[3]:
        kpi_card("Pedidos", money(total_pedido), f"{pedidos_qtd} pedidos filtrados", "#a78bfa")
    with cols[4]:
        kpi_card("A faturar", money(total_a_faturar), "Saldo em pedidos", "#ef4444")


def seller_performance(pedidos: pd.DataFrame, vendas: pd.DataFrame, metas: pd.DataFrame) -> pd.DataFrame:
    vendas_validas = vendas[vendas["Venda Valida"]]
    vend = vendas_validas.groupby("Vendedor", as_index=False).agg(
        Vendas=("Valor Total Liquido", "sum"),
        Notas=("Nota Fiscal", "count"),
    )
    ped = pedidos.groupby("Vendedor", as_index=False).agg(
        Pedidos=("Pedido", "count"),
        Valor_Pedidos=("Valor do Pedido", "sum"),
        A_Faturar=("Valor A Faturar", "sum"),
    )
    df = pd.merge(vend, ped, on="Vendedor", how="outer").fillna(0)
    df = pd.merge(df, metas, on="Vendedor", how="left").fillna({"Meta": 0})
    df["% Meta"] = df.apply(lambda r: r["Vendas"] / r["Meta"] if r["Meta"] else 0.0, axis=1)
    return df.sort_values("Vendas", ascending=False)


def render_tables(pedidos: pd.DataFrame, vendas: pd.DataFrame, metas: pd.DataFrame) -> None:
    perf = seller_performance(pedidos, vendas, metas)
    display_perf = perf.copy()
    for col in ["Vendas", "Valor_Pedidos", "A_Faturar", "Meta"]:
        display_perf[col] = display_perf[col].map(money)
    display_perf["% Meta"] = display_perf["% Meta"].map(pct)
    display_perf = display_perf.rename(columns={"Valor_Pedidos": "Valor Pedidos", "A_Faturar": "A Faturar"})

    st.subheader("Desempenho por vendedor")
    st.dataframe(display_perf, use_container_width=True, hide_index=True)

    if not perf.empty:
        fig = px.bar(perf, x="Vendedor", y=["Vendas", "Meta"], barmode="group", title="Vendas x Meta por vendedor")
        fig.update_layout(legend_title_text="", yaxis_title="Valor", xaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

    vendas_validas = vendas[vendas["Venda Valida"]].copy()
    diario = vendas_validas.groupby("Data", as_index=False)["Valor Total Liquido"].sum().sort_values("Data")
    if not diario.empty:
        st.subheader("Vendas diarias")
        fig = px.line(diario, x="Data", y="Valor Total Liquido", markers=True)
        fig.update_layout(yaxis_title="Valor vendido", xaxis_title="Data")
        st.plotly_chart(fig, use_container_width=True)

    tab_vendas, tab_pedidos = st.tabs(["Vendas", "Pedidos"])
    with tab_vendas:
        cols = ["Data", "Nota Fiscal", "Cliente", "Vendedor", "Valor Total Liquido", "Cancelada?"]
        view = vendas[cols].sort_values("Data", ascending=False).copy()
        view["Data"] = view["Data"].dt.strftime("%d/%m/%Y")
        view["Valor Total Liquido"] = view["Valor Total Liquido"].map(money)
        st.dataframe(view, use_container_width=True, hide_index=True)
    with tab_pedidos:
        cols = ["Data", "Pedido", "Cliente", "Vendedor", "Valor do Pedido", "Valor A Faturar", "Situacao Atendimento"]
        view = pedidos[cols].sort_values("Data", ascending=False).copy()
        view["Data"] = view["Data"].dt.strftime("%d/%m/%Y")
        view["Valor do Pedido"] = view["Valor do Pedido"].map(money)
        view["Valor A Faturar"] = view["Valor A Faturar"].map(money)
        st.dataframe(view, use_container_width=True, hide_index=True)


def main() -> None:
    style_app()
    with st.sidebar:
        st.markdown("## Arquivos")
        pedido_source, pedido_name = source_selector("Pedido", PEDIDO_XLS, "pedido_upload")
        venda_source, venda_name = source_selector("Venda", VENDA_XLS, "venda_upload")

    if LOGO_PATH.exists():
        c_logo, c_title = st.columns([0.12, 0.88], vertical_alignment="center")
        c_logo.markdown(logo_html(LOGO_PATH), unsafe_allow_html=True)
        c_title.markdown(
            """
            <div class="fardin-header">
              <h1>Dashboard Comercial - Fardin</h1>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div class="fardin-header">
              <h1>Dashboard Comercial - Fardin</h1>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if pedido_source is None or venda_source is None:
        st.warning("Informe os arquivos Pedido.XLS e Venda.XLS para carregar o dashboard.")
        st.stop()

    try:
        pedidos = load_pedidos(pedido_source)
        vendas = load_vendas(venda_source)
    except Exception as exc:
        st.error("Nao foi possivel carregar os arquivos selecionados.")
        st.exception(exc)
        st.stop()

    pedidos_f, vendas_f = apply_filters(pedidos, vendas)
    vendedores_filtrados = sorted(set(pedidos_f["Vendedor"]).union(vendas_f["Vendedor"]) - {""})
    total_vendido_filtrado = float(vendas_f[vendas_f["Venda Valida"]]["Valor Total Liquido"].sum())
    meta_total, metas = build_meta_editor(vendedores_filtrados, total_vendido_filtrado)

    render_kpis(pedidos_f, vendas_f, meta_total)
    render_tables(pedidos_f, vendas_f, metas)


if __name__ == "__main__":
    main()

