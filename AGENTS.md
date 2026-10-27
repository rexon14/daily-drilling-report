## Cursor Cloud specific instructions

### Overview

This is a Python/Streamlit multi-page web application for processing daily drilling reports (petroleum/oil & gas domain). It transforms plain-text and Excel drilling reports into a standardized 19-column format.

### Running the app

```bash
streamlit run app.py --server.enableCORS false --server.enableXsrfProtection false --server.headless true --server.port 8501
```

The app serves on port **8501**. Multi-page sub-apps (`pages/region-1.py`, `pages/region-2.py`) are auto-discovered by Streamlit.

### Key caveats

- **PATH**: The `streamlit` binary installs to `~/.local/bin`. Ensure `export PATH="$HOME/.local/bin:$PATH"` is in effect before running commands.
- **No external services required**: No databases, Docker, or message queues. This is a standalone Python app.
- **Zone processors**: When a report is saved, the main app invokes zone-specific `app.py` scripts (e.g. `Zone 7/app.py`) via `subprocess.run()`. These expect report text in a domain-specific format; arbitrary text will trigger `ValueError: No well data found`.
- **Directories with spaces**: Several source directories contain spaces (`Region 1/`, `Region 2/`, `Region 5/`, `Zone 7/`, etc.). Always quote paths.
- **Lint**: Run `ruff check .` from the workspace root. There are ~36 pre-existing lint warnings (unused imports/variables). Use `ruff check . --select E9,F63,F7,F82` for syntax-error-only checks.
- **No automated test suite**: The repository does not include unit or integration tests.
