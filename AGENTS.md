# Repository Guidelines

## Project Structure & Module Organization

This repository contains a Streamlit dashboard for Doces Fardin commercial reporting.

- `app.py` is the Streamlit entry point. It reloads and delegates execution to `app_fardin.main()`.
- `app_fardin.py` contains the dashboard logic, Excel loading, filters, manual goals, KPIs, charts, and tables.
- `requirements.txt` lists Python runtime dependencies, including `xlrd` for legacy `.XLS` files.
- `Pedido.XLS` and `Venda.XLS` are the current local default data files.
- `logo fardin.jpg` and `Logo Alfa.jpg` are UI assets.
- `backup-codigo-*` folders are snapshots of prior code and should not be edited during feature work.

Keep new dashboard logic in focused helper functions. Avoid reintroducing dependencies on the old summary/evaluation workbooks.

## Build, Test, and Development Commands

Create and activate a virtual environment before installing dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Run the app locally with:

```powershell
streamlit run app.py
```

Use this command for manual validation after changes to Excel parsing, filters, meta inputs, KPIs, or charts.

## Coding Style & Naming Conventions

Use Python 3 style with 4-space indentation. Constants use `UPPER_SNAKE_CASE`; functions and variables use `snake_case`. Keep formatting helpers deterministic and side-effect free. Preserve Brazilian currency, date, and percentage formatting conventions in user-facing values.

## Testing Guidelines

There is no automated test suite yet. Validate manually with `Pedido.XLS` and `Venda.XLS`. When adding tests, prefer `pytest`, name files `test_*.py`, and focus first on pure helpers for money parsing, date parsing, filtering, and goal calculations.

## Commit & Pull Request Guidelines

Recent commits use short, imperative Portuguese summaries, for example `Ajustar layout desktop no Streamlit Cloud`. Keep commit subjects concise and focused on one change.

Pull requests should include a brief description, manual test notes, affected data files or upload paths, and screenshots when the visual layout changes.

## Security & Configuration Tips

Do not commit credentials, private spreadsheets, or exported customer data. Treat uploaded Excel files as sensitive business data.
