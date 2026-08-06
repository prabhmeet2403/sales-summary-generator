# Deploying to Render

## What this is
This repository contains several distinct pieces:
- **The calculation engine** (`config.py`, `excel_reader.py`, `aggregator.py`, `monthly_view.py`, `comment_mapper.py`, `historical_lookup.py`, `summary_writer.py`, `sheet_copy.py`, `validator.py`, `row_validator.py`, `column_autofit.py`) -- untouched, no deployment-related change affects this.
- **The CLI entry point** (`main.py`) -- for running the generator from a terminal / batch job. Not used by the Render deployment.
- **The desktop GUI** (`gui_main.py`, `SalesForecastGUI.spec`, `build_exe.bat`) -- a separate, Windows-oriented distribution (a standalone .exe via PyInstaller). Not used by, and not needed for, the Render deployment. Safe to leave in the repo; Render never touches these files.
- **The Streamlit web app** (`app.py`, `streamlit_bridge.py`, `components/`, `ai/`) -- **this is what gets deployed to Render.**

## One security issue found and fixed before this deployment setup
The uploaded project included a real `.env` file containing what appears to be a live AWS Bedrock bearer token. **This has been removed.** In its place there is now a safe `.env.example` template with placeholder values. The real secret was never written into any file in this delivery -- if that token was ever real, it should be rotated in AWS regardless, since it existed in a shared archive.

Going forward: never commit a real `.env`. `.gitignore` already excludes it. Secrets are set as real environment variables in Render's dashboard instead (see below).

## Entry point
**`app.py`** is the correct Streamlit entry point (confirmed both by its own docstring, `"""Run with: streamlit run app.py"""`, and by actually starting it in a clean virtual environment during this review -- it launched successfully). `main.py` and `gui_main.py` are separate entry points for other, non-web uses and are not involved in this deployment.

## Files added for this deployment
| File | Purpose |
|---|---|
| `render.yaml` | Render Blueprint -- defines the service, build/start commands, Python version, and the three secret env vars (as empty placeholders you fill in via the dashboard) |
| `.env.example` | Safe template for local development; replaces the removed real `.env` |
| `Procfile` | Fallback start-command definition (Render primarily uses `render.yaml`; this covers other platforms/tools that read a Procfile) |
| `runtime.txt` | Pins the Python version as a fallback to `render.yaml`'s `PYTHON_VERSION` |

`.gitignore` was extended (nothing removed) to also exclude `.streamlit/secrets.toml`, virtual environment folders, and common OS/editor clutter.

## Deploying

### Option A -- Blueprint (recommended, uses render.yaml)
1. Push this repository to GitHub/GitLab.
2. In the Render dashboard: **New -> Blueprint**, point it at the repo. Render reads `render.yaml` automatically.
3. Render will create the web service with the build/start commands already configured. It will prompt you to fill in the three secret values (`BEDROCK_REGION`, `BEDROCK_MODEL_ID`, `AWS_BEARER_TOKEN_BEDROCK`) since they're marked `sync: false` -- they are never stored in the repo or in `render.yaml` itself.
4. Deploy.

### Option B -- Manual web service
1. In the Render dashboard: **New -> Web Service**, connect the repo.
2. Runtime: Python 3.
3. Build command: `pip install -r requirements.txt`
4. Start command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
5. Under **Environment**, add `BEDROCK_REGION`, `BEDROCK_MODEL_ID`, `AWS_BEARER_TOKEN_BEDROCK` with your real values.
6. Deploy.

## Verified before delivery
- `requirements.txt` installed cleanly into a fresh virtual environment, and `streamlit run app.py` was actually started (not just assumed) in that environment -- confirmed the Uvicorn server came up with no import errors.
- No hardcoded Windows-only paths (`C:\...`, literal backslashes) anywhere in the codebase -- all paths use `pathlib.Path`, confirmed via `PROJECT_ROOT = Path(__file__).resolve().parent` in `config.py` and the same pattern throughout.
- Uploaded-file handling (`app.py`) writes into `tempfile.mkdtemp()` (OS-appropriate temp directory, not a fixed path) and cleans up the previous session's temp directory (`shutil.rmtree`) each time a new file is uploaded.
- `app.py` starts and runs correctly with `.env` entirely absent (confirmed directly) -- `load_dotenv()` is a no-op if the file doesn't exist, and Render's real environment variables take its place in production.

## Not changed
No business logic, no Excel-calculation logic, no file under the calculation engine listed above. Every change in this delivery is deployment configuration only: `render.yaml`, `Procfile`, `runtime.txt`, `.env.example`, `.gitignore`, and the removal of the one file that should never have been included (`.env`).
