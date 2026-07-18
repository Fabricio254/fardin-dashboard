# Repository Guidelines

## Project Structure & Module Organization

This repository contains a Streamlit dashboard for Doces Fardin commercial reporting.

- `app.py` is the Streamlit entry point. It reloads and delegates execution to `app_fardin.main()`.
- `app_fardin.py` contains the dashboard logic, data loading, formatting helpers, access rules, and UI rendering.
- `requirements.txt` lists Python runtime dependencies.
- `logo fardin.jpg` and `Logo Alfa.jpg` are UI assets used by the dashboard.
- Expected Excel workbooks are loaded from the repository root when present, with fallback names for accented filenames.

Keep new source code close to the current structure unless the module becomes difficult to navigate. Prefer small helper functions in `app_fardin.py` before introducing additional modules.

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

Use this command for manual validation after changes, especially when touching file upload behavior, Excel parsing, authentication/demo access, or chart rendering.

## Coding Style & Naming Conventions

Use Python 3 style with 4-space indentation. Follow the existing pragmatic style: module-level constants in `UPPER_SNAKE_CASE`, helper functions in `snake_case`, and clear Portuguese labels where they are part of the user-facing dashboard. Keep formatting helpers deterministic and side-effect free.

Avoid broad refactors when changing dashboard behavior. Preserve existing Brazilian number, date, currency, and percentage formatting conventions.

## Testing Guidelines

There is no automated test suite in this repository yet. For now, validate changes manually with `streamlit run app.py` using representative Excel files. When adding tests, prefer `pytest`, name files `test_*.py`, and focus first on pure helpers such as money/date formatting, numeric conversion, workbook loading, and access-rule calculations.

## Commit & Pull Request Guidelines

Recent commits use short, imperative Portuguese summaries, for example `Ajustar layout desktop no Streamlit Cloud`. Keep commit subjects concise and focused on one change.

Pull requests should include a brief description, manual test notes, affected data files or upload paths, and screenshots when the visual layout changes.

## Security & Configuration Tips

Do not commit real credentials, private spreadsheets, or exported customer data. Treat password and demo-access changes as sensitive, and document any expected Streamlit Cloud environment differences in the PR notes.
