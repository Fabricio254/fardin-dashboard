# Repository Guidelines – Doces Fardin Dashboard

## Project Structure & Module Organization

This repository contains a **Streamlit dashboard** for Doces Fardin commercial reporting. It loads sales/order data from Excel files and renders KPIs, filters, manual goals, and performance tables.

### Core Files
| File | Purpose |
|------|---------|
| `app.py` | Entry point — reloads and delegates to `app_fardin.main()` |
| `app_fardin.py` | Main dashboard logic: loading, filtering, KPIs, tables, charts |
| `requirements.txt` | Python dependencies (Streamlit, Pandas, Plotly, xlrd for `.XLS`) |
| `Pedido.XLS`, `Venda.XLS` | Default local data files (git-ignored on client machines) |
| `logo fardin.jpg`, `Logo Alfa.jpg` | UI branding assets |
| `.streamlit/config.toml` | Streamlit configuration |

### Data Flow
```
Pedido.XLS / Venda.XLS (uploaded or default)
     ↓
load_pedidos() / load_vendas() → parse dates, money, validate rows
     ↓
apply_filters() → filter by Vendedor, Período, etc.
     ↓
render_kpis() → calculate totals, % atingimento, ticket, crescimento
render_tables() → seller performance, order details
```

### Architecture Notes
- **No persistent database**: All data flows from Excel files; filters and metas are session-scoped (except manual meta input via `build_meta_editor()`)
- **Helper functions are pure**: `money()`, `pct()`, date/money parsing should remain deterministic and side-effect-free
- **Backup folders** (`backup-codigo-*`): snapshots of prior code — do not edit during feature work

## Build, Test, and Development Commands

### Initial Setup
Create and activate a virtual environment before installing dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Run Locally
```powershell
streamlit run app.py
```
- Dashboard opens at `http://localhost:8501`
- File selector sidebar allows uploading custom `Pedido.XLS` and `Venda.XLS`
- Default files loaded from workspace root if not uploaded

### Manual Validation Checklist
Test after changes to:
- **Excel parsing** → ensure `load_pedidos()` / `load_vendas()` handle date/money formats correctly
- **Filters** → verify `apply_filters()` correctly segments by Vendedor, Período
- **Meta input** → check `build_meta_editor()` saves/calculates goals correctly
- **KPIs** → validate `render_kpis()` calculations (%, ticket, crescimento)
- **Charts/tables** → visual inspection of `render_tables()` output

**No automated test suite yet.** When adding tests, prefer `pytest` and name files `test_*.py`.

## Coding Style & Naming Conventions

Use Python 3 style with 4-space indentation. Constants use `UPPER_SNAKE_CASE`; functions and variables use `snake_case`. Keep formatting helpers deterministic and side-effect free. Preserve Brazilian currency, date, and percentage formatting conventions in user-facing values.

## Key Functions & Common Modification Patterns

### Data Loading & Parsing
- `load_pedidos(source)` → reads Pedido.XLS, parses dates/money, validates row structure
- `load_vendas(source)` → reads Venda.XLS, filters "Venda Valida", calculates totals
- `parse_money_series()` / `parse_date_series()` → pure functions, safe to modify for format changes
- **When modifying Excel structure:** update both parsing functions AND `read_xls_table()` row validation

### Filtering & Aggregation
- `apply_filters()` → returns (pedidos_filtered, vendas_filtered) tuples
- Filters are **session-scoped**: no persistence except manual meta input
- **When adding new filters:** update sidebar source_selector() and apply_filters() logic

### KPI Calculations
- `render_kpis()` → computes meta%, ticket, crescimento, seller totals
- **When modifying calculations:** ensure `build_meta_editor()` is called first (depends on total_vendido_filtrado)

### UI Components
- `kpi_card()` → reusable KPI display component (label, value, note, accent color)
- `logo_html()` / `style_app()` → branding and CSS customization
- **Streamlit patterns:** `st.sidebar`, `st.columns()` for layout, `st.session_state` for metas

## Testing Guidelines

Focus first on pure helpers for money parsing, date parsing, filtering, and goal calculations. When adding tests, use `pytest` and name files `test_*.py`.

## Commit & Pull Request Guidelines

Recent commits use short, imperative Portuguese summaries, for example `Ajustar layout desktop no Streamlit Cloud`. Keep commit subjects concise and focused on one change.

Pull requests should include a brief description, manual test notes, affected data files or upload paths, and screenshots when the visual layout changes.

## Security & Configuration Tips

Do not commit credentials, private spreadsheets, or exported customer data. Treat uploaded Excel files as sensitive business data. Files in `.gitignore`: `*.xlsx`, `*.xls`, `*.XLS`, `.env`, `venv/`, `backup-codigo-*/`.

## Git-Ignored Files & Local Data

The following are **never committed** and expected to exist on client machines:
- `Pedido.XLS`, `Venda.XLS` (default data files)
- `.venv/` (Python virtual environment)
- `Thumbs.db`, `.DS_Store` (OS artifacts)
- `.streamlit/secrets.toml` (sensitive config)

When developing features that reference these files, ensure fallback handling via `source_selector()` and clear error messages.
