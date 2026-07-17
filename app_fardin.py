from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
PEDIDO_XLS = BASE_DIR / "Pedido.XLS"
VENDA_XLS = BASE_DIR / "Venda.XLS"
LOGO_PATH = BASE_DIR / "logo fardin.jpg"

WorkbookSource = str | Path | bytes


st.set_page_config(
    page_title="Dashboard Comercial - Fardin",
    page_icon="F",
    layout="wide",
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
    upload = st.sidebar.file_uploader(label, type=["xls", "xlsx"], key=key)
    if upload is not None:
        return upload.getvalue(), upload.name
    if default_path.exists():
        return default_path, default_path.name
    return None, "Aguardando arquivo"


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
    date_range = st.sidebar.date_input("Periodo de venda", value=(min_date, max_date), min_value=min_date, max_value=max_date)
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
    else:
        start = end = date_range

    vendedores = sorted(set(pedidos["Vendedor"]).union(vendas["Vendedor"]) - {""})
    selected = st.sidebar.multiselect("Vendedor", vendedores, default=vendedores)

    start_ts = pd.Timestamp(start)
    end_ts = pd.Timestamp(end)
    pedidos_f = pedidos[pedidos["Data"].between(start_ts, end_ts)].copy()
    vendas_f = vendas[vendas["Data"].between(start_ts, end_ts)].copy()
    if selected:
        pedidos_f = pedidos_f[pedidos_f["Vendedor"].isin(selected)]
        vendas_f = vendas_f[vendas_f["Vendedor"].isin(selected)]
    return pedidos_f, vendas_f


def render_kpis(pedidos: pd.DataFrame, vendas: pd.DataFrame, meta_total: float) -> None:
    vendas_validas = vendas[vendas["Venda Valida"]]
    total_vendido = float(vendas_validas["Valor Total Liquido"].sum())
    total_pedido = float(pedidos["Valor do Pedido"].sum())
    total_a_faturar = float(pedidos["Valor A Faturar"].sum())
    atingimento = total_vendido / meta_total if meta_total else 0.0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Vendas validas", money(total_vendido))
    c2.metric("Meta total", money(meta_total))
    c3.metric("Atingimento", pct(atingimento))
    c4.metric("Pedidos", money(total_pedido))
    c5.metric("A faturar", money(total_a_faturar))


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
    with st.sidebar:
        st.markdown("## Arquivos")
        pedido_source, pedido_name = source_selector("Pedido", PEDIDO_XLS, "pedido_upload")
        venda_source, venda_name = source_selector("Venda", VENDA_XLS, "venda_upload")
        st.caption(f"Pedido: {pedido_name}")
        st.caption(f"Venda: {venda_name}")

    if LOGO_PATH.exists():
        c_logo, c_title = st.columns([0.12, 0.88], vertical_alignment="center")
        c_logo.image(str(LOGO_PATH), width=90)
        c_title.title("Dashboard Comercial - Fardin")
    else:
        st.title("Dashboard Comercial - Fardin")

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

