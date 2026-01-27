from __future__ import annotations

import json
import tempfile
from dataclasses import asdict
from pathlib import Path

import streamlit as st

from lm5148_tool.export_results_xlsx import build_results_xlsx_bytes
from lm5148_tool.lm5148_design_tool import (
    DesignInputs,
    run_design,
    eq31_l_required,
)
from lm5148_tool.quickstart_excel_com import fill_quickstart_excel


def _default_template_path() -> Path:
    # training/ is at repo root
    return Path(__file__).resolve().parent.parent / "training" / "LM5148_LM25148_quickstart_calculator_A4.xlsm"


st.set_page_config(page_title="LM5148 One-Place Tool", layout="wide")

st.title("LM5148 design helper + exports")
st.caption("Single place to calculate and export: JSON + results .xlsx + filled TI quickstart .xlsm/.xlsx (Excel required).")

with st.sidebar:
    st.header("Inputs")
    vin_min = st.number_input("VIN min [V]", value=10.0, step=0.5)
    vin_nom = st.number_input("VIN nominal [V]", value=12.0, step=0.5)
    vin_max = st.number_input("VIN max [V]", value=18.0, step=0.5)

    vout = st.number_input("VOUT [V]", value=5.0, step=0.1)
    iout = st.number_input("IOUT [A]", value=8.0, step=0.5)

    fsw_mhz = st.number_input("FSW [MHz]", value=2.1, step=0.1)
    fsw_hz = fsw_mhz * 1e6

    st.divider()
    st.subheader("Options")
    template_path = st.text_input("TI quickstart template (.xlsm)", value=str(_default_template_path()))
    excel_visible = st.checkbox("Show Excel while exporting", value=False)

# Minimal set of inputs for the existing Python design flow
inp = DesignInputs(
    vin_nom_v=vin_nom,
    vin_max_v=vin_max,
    vout_v=vout,
    iout_a=iout,
    fsw_hz=fsw_hz,
)

res = run_design(inp)

# Build a payload compatible with the existing webapp exporter
payload = {
    "meta": {"tool": "lm5148_streamlit_app", "version": 1},
    "inputs": {
        "vinMin": vin_min,
        "vinNom": vin_nom,
        "vinMax": vin_max,
        "vout": vout,
        "iout": iout,
        "fsw": fsw_hz,
    },
    "results": asdict(res),
}

colA, colB = st.columns([1, 1])

with colA:
    st.subheader("Key results")
    st.write(
        {
            "L required (Eq31) [H]": res.l_required_h,
            "ΔIL @ VINmax (Eq32) [A]": res.delta_il_vin_max_a,
            "IL peak @ VINmax (Eq32) [A]": res.il_peak_vin_max_a,
            "RSENSE (Eq34) [Ω]": res.rsense_ohm,
            "IL peak short (Eq35) [A]": res.il_peak_short_a,
            "Cout min load-off (Eq36) [F]": res.cout_load_off_f,
            "Vout ripple pp (Eq37) [V]": res.vout_ripple_pp_v,
            "Cin required (Eq40) [F]": res.cin_required_f,
        }
    )

with colB:
    st.subheader("Exports")

    st.download_button(
        "Download JSON (for tooling)",
        data=json.dumps(payload, indent=2).encode("utf-8"),
        file_name="lm5148_design.json",
        mime="application/json",
        use_container_width=True,
    )

    xlsx_bytes = build_results_xlsx_bytes(inputs=payload["inputs"], results=payload["results"])
    st.download_button(
        "Download results.xlsx (standalone)",
        data=xlsx_bytes,
        file_name="lm5148_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("**TI Quickstart exports (Excel required on Windows)**")
    if st.button("Build filled quickstart .xlsm/.xlsx", use_container_width=True):
        try:
            with tempfile.TemporaryDirectory() as td:
                td_path = Path(td)
                json_path = td_path / "lm5148_design.json"
                json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

                out_xlsm = td_path / "LM5148_quickstart_filled.xlsm"
                out_xlsx = td_path / "LM5148_quickstart_filled.xlsx"

                fill_quickstart_excel(
                    json_path=json_path,
                    template_path=Path(template_path),
                    out_xlsm=out_xlsm,
                    out_xlsx=out_xlsx,
                    visible=excel_visible,
                )

                st.success("Built quickstart exports.")
                st.download_button(
                    "Download filled .xlsm",
                    data=out_xlsm.read_bytes(),
                    file_name=out_xlsm.name,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    use_container_width=True,
                )
                st.download_button(
                    "Download filled .xlsx",
                    data=out_xlsx.read_bytes(),
                    file_name=out_xlsx.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        except Exception as e:
            st.error(str(e))

st.divider()
st.subheader("Notes")
st.write(
    "This Streamlit app is meant to unify exports. It currently focuses on the same core inputs the TI quickstart sheet accepts "
    "(VIN/VOUT/IOUT/FSW). The static web calculator remains the most detailed step-by-step view."
)
