# LM5148 Tooling (one place)

This folder contains the LM5148 design helpers and export tooling.

## Option A — Run the interactive web calculator

See [lm5148_webapp/README.md](lm5148_webapp/README.md).

## Option B — Run the Python app (Streamlit)

### 1) Use the supported Python environment

Streamlit does **not** support Python 3.14.
Use the Python 3.10 venv in this workspace:

- Interpreter: `.venv310` (recommended)

### 2) Install dependencies

From the repo root:

- `& "m:/Amit Medina/Schematic and Layout Design Course/.venv310/Scripts/python.exe" -m pip install -r "Schematic-and-Layout-Design-Course-repo/lm5148_tool/requirements.txt"`

### 3) Run Streamlit

- `& "m:/Amit Medina/Schematic and Layout Design Course/.venv310/Scripts/python.exe" -m streamlit run "Schematic-and-Layout-Design-Course-repo/lm5148_tool/lm5148_streamlit_app.py"`

## Export to TI Quickstart (.xlsm)

The **best fidelity** export (preserves macros/shapes and forces recalculation) requires:
- Windows + Microsoft Excel
- Optional Python dependencies:
  - `pip install -r requirements-excel-automation.txt`

Then the Streamlit app can generate both:
- Filled `.xlsm`
- Exported `.xlsx`

Or run the CLI directly:

- `python lm5148_tool/quickstart_excel_com.py --json lm5148_tool/lm5148_design.json --template training/LM5148_LM25148_quickstart_calculator_A4.xlsm --out-xlsm lm5148_tool/LM5148_quickstart_filled.xlsm --out-xlsx lm5148_tool/LM5148_quickstart_filled.xlsx`
