# LM5148 Web Calculator (pages 36–39)

## Open from anywhere (GitHub Pages)

If you enabled **GitHub Pages** for this repo, the calculator is available at:

https://amitmedina.github.io/Schematic-and-Layout-Design-Course/lm5148_tool/lm5148_webapp/
 
- `https://<your-user>.github.io/<repo-name>/lm5148_tool/lm5148_webapp/`

For example, if your GitHub user is `amitmedina` and the repo name is `Schematic-and-Layout-Design-Course`, the direct URL is:

- `https://amitmedina.github.io/Schematic-and-Layout-Design-Course/lm5148_tool/lm5148_webapp/`

To enable it: GitHub → **Settings → Pages** → **Build and deployment** → Deploy from a branch → Branch: `main` → Folder: `/ (root)`.

## Open on your phone (same Wi‑Fi)

1. From the repo root, start a local server:
	- `py -m http.server 8000 --bind 0.0.0.0`
2. On your PC, find your IPv4 address (Wi‑Fi adapter):
	- `ipconfig`
3. On your phone (connected to the same Wi‑Fi), open:
	- `http://<YOUR_PC_IP>:8000/lm5148_tool/lm5148_webapp/`

If it doesn’t load, allow Python through Windows Firewall (Private network) or try a different port.

Static HTML app that follows the LM5148 datasheet design flow (pages 36–39) and computes equations (31–45) with inputs placed next to each equation.

## Run

Option A (simplest):
- Open `lm5148_webapp/index.html` in a browser.

Option B (recommended in VS Code):
- Use the **Live Server** extension and open `lm5148_webapp/index.html`.

Option C (no extensions):
- From the workspace root run: `python -m http.server 8000`
- Then open: `http://localhost:8000/lm5148_tool/lm5148_webapp/`

## Notes

- The UI structure and equation numbering are aligned to the datasheet, but the written step text is paraphrased.
- Computations are a learning aid; validate the final design against the full datasheet and your requirements.
- Results display in engineering units (p/n/µ/m/k/M/...).

## Export to TI Quickstart (.xlsm)

1. In the app, click **Download JSON** to save `lm5148_design.json`.
2. Recommended (best fidelity): run the PowerShell Excel automation exporter (preserves shapes/macros and forces recalculation):

	`powershell -NoProfile -ExecutionPolicy Bypass -File "m:/Amit Medina/Schematic and Layout Design Course/populate_quickstart_calculator_excel.ps1" -JsonPath "m:/Amit Medina/Schematic and Layout Design Course/lm5148_design.json" -TemplatePath "m:/Amit Medina/Schematic and Layout Design Course/Schematic-and-Layout-Design-Course-repo/training/LM5148_LM25148_quickstart_calculator_A4.xlsm" -OutXlsm "m:/Amit Medina/Schematic and Layout Design Course/LM5148_quickstart_filled_excel.xlsm" -OutXlsx "m:/Amit Medina/Schematic and Layout Design Course/LM5148_quickstart_filled_excel.xlsx"`

3. Fallback (no Excel required): run the Python exporter (writes inputs but may drop some shapes/drawings in complex templates):

	`& "m:/Amit Medina/Schematic and Layout Design Course/.venv/Scripts/python.exe" "m:/Amit Medina/Schematic and Layout Design Course/Schematic-and-Layout-Design-Course-repo/lm5148_tool/populate_quickstart_calculator.py" --json "m:/Amit Medina/Schematic and Layout Design Course/Schematic-and-Layout-Design-Course-repo/lm5148_tool/lm5148_design.json" --out "m:/Amit Medina/Schematic and Layout Design Course/Schematic-and-Layout-Design-Course-repo/lm5148_tool/LM5148_quickstart_filled_v2.xlsm" --out-xlsx "m:/Amit Medina/Schematic and Layout Design Course/Schematic-and-Layout-Design-Course-repo/lm5148_tool/LM5148_quickstart_filled_v2.xlsx"`

## One-place Python app (Streamlit)

If you prefer to have **calculations + exports** in one place (instead of a static web page + separate scripts), run the Streamlit app:

- Install deps (from `lm5148_tool/`): `pip install -r requirements.txt`
- Run: `streamlit run lm5148_streamlit_app.py`

It supports:
- Download `lm5148_design.json`
- Download a standalone `lm5148_results.xlsx`
- Build and download filled TI quickstart outputs:
	- `.xlsm` (macro-enabled, preserves shapes/macros)
	- `.xlsx`

Note: the `.xlsm` export path requires **Windows + Microsoft Excel** and the `pywin32` package.

Install the optional Excel automation deps with:

`pip install -r lm5148_tool/requirements-excel-automation.txt`

## Python COM exporter (no PowerShell)

If you want the best-fidelity `.xlsm` export but prefer Python over PowerShell, use:

`python lm5148_tool/quickstart_excel_com.py --json lm5148_tool/lm5148_design.json --template training/LM5148_LM25148_quickstart_calculator_A4.xlsm --out-xlsm lm5148_tool/LM5148_quickstart_filled_excel.xlsm --out-xlsx lm5148_tool/LM5148_quickstart_filled_excel.xlsx`

## Publish (GitHub Pages)

If you want to open the calculator from your phone (or anywhere), publish it via **GitHub Pages**:

1. Push this repo to GitHub.
2. In GitHub: **Settings → Pages**.
3. Under **Build and deployment** choose:
	- **Source**: Deploy from a branch
	- **Branch**: `main`
	- **Folder**: `/ (root)`
4. After it deploys, open the site:
	- Main landing page: `https://<your-user>.github.io/<repo-name>/`
	- Direct app URL: `https://<your-user>.github.io/<repo-name>/lm5148_tool/lm5148_webapp/`

Note: VS Code’s *internal preview/webview* often blocks external scripts (like MathJax from a CDN), so the page can look “broken” there. Use a real browser (Chrome/Edge) or serve it over `http://` (Live Server / `python -m http.server`) and it will render correctly.
