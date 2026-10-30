# Excel VLOOKUP Studio

Fills STATUS, COMMENTS (or any last columns) in every sheet of the latest workbook from an older one,
matched on the key columns you choose. It warns when counts differ and removes `#N/A`.
The download always contains every sheet of the latest file.

> **Patient data (PHI):** files are processed in memory only and are never written to disk, but on a hosted
> service they are uploaded to that service's servers. Streamlit Community Cloud does not sign a HIPAA BAA.
> On a cloud deployment, make the app **private** (invited viewers only), turn on the password, and follow your
> organisation's rules for PHI. For PHI, the safest choice is running it locally or on a company server.

## Option 1: Streamlit Community Cloud

1. Sign in at [share.streamlit.io](https://share.streamlit.io) with the GitHub account that owns this repository.
2. Click **Create app**, choose to deploy from a GitHub repository, and fill in:
   - Repository: this repository
   - Branch: `main`
   - Main file path: `app.py`
3. Open **Advanced settings**:
   - **Python version:** 3.12 or 3.13
   - **Secrets:** paste the line below with your own password
     ```toml
     password = "choose-a-strong-password"
     ```
4. Click **Deploy**. The first start takes a few minutes while the packages install.
5. In the app, go to **Share** and restrict who can view it to the people you invite by email.

To update the app, push to `main`; Streamlit Cloud redeploys on its own.
Uploads are limited to 200 MB per file (`.streamlit/config.toml`), because cloud apps have limited memory.

## Option 2: on each PC

1. Install Python 3.11 or newer from python.org. Tick **Add python.exe to PATH** during the install.
2. Copy this folder to the PC.
3. Double-click **Start VLOOKUP Studio.bat**. It installs the packages the first time, then opens the browser.

## Option 3: one server for the whole team (internal network)

With Docker:

```bash
docker build -t vlookup-studio .
docker run -d --name vlookup-studio --restart unless-stopped -p 8501:8501 -e VLOOKUP_PASSWORD=choose-a-password vlookup-studio
```

Without Docker:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

People open `http://<server-name>:8501`. Health check URL: `http://<server-name>:8501/_stcore/health`.
Put it behind your company's HTTPS reverse proxy (IIS, nginx), which must allow WebSockets (`/_stcore/stream`).

## Password

The app asks for a password when one is set. Use `password` in Streamlit secrets (Cloud: app settings >
Secrets; locally: copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml`) or the `VLOOKUP_PASSWORD`
environment variable. Without either one, the app opens with no sign-in.

## Checks before each release

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q tests
```

The tests use made-up data only. They cover matching, the count check, `#N/A` removal, fallback keys,
duplicate rules, the UNIQUE KEY, keeping all sheets, merged cells and damaged or protected files. They also
check that the page and the password screen load.

## Error handling

| Situation | What the user sees |
|---|---|
| CSV / .xls renamed to .xlsx, damaged file, empty file | A plain message explaining how to re-save it in Excel |
| Password-protected workbook | A message asking to remove the password and save again |
| Header row number past the data | A warning to check the header row, and the sheet is skipped |
| UNIQUE KEY named like a key or fill column | The sheet is blocked until it is renamed |
| UNIQUE KEY column already has data | A warning that it will be replaced |
| Merged cells where values are written | Unmerged automatically, then filled |
| Text over Excel's 32,767-character limit / invalid characters | Cut / removed so the file still opens |
| Charts or pictures in the workbook | A warning that they may be simplified in the new file |
| Very large sheet (over 200,000 cells) | The preview is shown without colours instead of failing |
| Anything unexpected | A friendly message plus "Technical details", logged on the server. Nothing is downloaded half-done |

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit screens (entry point) |
| `vlookup_engine.py` | Reading, matching, count check, writing the new workbook |
| `requirements.txt` | Exact package versions (used by Streamlit Cloud, Docker and the .bat) |
| `.streamlit/config.toml` | Theme, upload limit, usage statistics off |
| `Dockerfile` | Server image (runs as a non-root user, with a health check) |
| `tests/` | Automated checks |
