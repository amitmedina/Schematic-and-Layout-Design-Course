from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class QuickstartInputs:
    vin_min_v: float
    vin_nom_v: float
    vin_max_v: float
    vout_v: float
    iout_a: float
    fsw_hz: float


def _load_inputs_from_webapp_json(json_path: Path) -> QuickstartInputs:
    payload = json.loads(json_path.read_text(encoding="utf-8"))
    inputs: dict[str, Any] = payload.get("inputs") or {}

    vin_nom = float(inputs["vinNom"])
    vin_min = float(inputs.get("vinMin", vin_nom))
    vin_max = float(inputs["vinMax"])

    return QuickstartInputs(
        vin_min_v=vin_min,
        vin_nom_v=vin_nom,
        vin_max_v=vin_max,
        vout_v=float(inputs["vout"]),
        iout_a=float(inputs["iout"]),
        fsw_hz=float(inputs["fsw"]),
    )


def fill_quickstart_excel(
    *,
    json_path: Path,
    template_path: Path,
    out_xlsm: Path,
    out_xlsx: Path | None = None,
    visible: bool = False,
) -> None:
    """Fill TI's macro-enabled LM5148 quickstart calculator using real Excel.

    Notes:
    - Requires Windows + Microsoft Excel installed.
    - Preserves macros/shapes and forces recalculation (best fidelity).
    """

    try:
        import win32com.client  # type: ignore
    except Exception as e:  # pragma: no cover
        raise RuntimeError(
            "pywin32 is required for Excel export. Install with: pip install pywin32"
        ) from e

    if template_path.name.startswith("~$"):
        raise ValueError(
            "Template path points at an Excel temp/lock file (~$...). Use the real .xlsm template."
        )

    inp = _load_inputs_from_webapp_json(json_path)
    fsw_khz = inp.fsw_hz / 1000.0

    out_xlsm.parent.mkdir(parents=True, exist_ok=True)
    if out_xlsx is not None:
        out_xlsx.parent.mkdir(parents=True, exist_ok=True)

    excel = None
    wb = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = bool(visible)
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(str(template_path), 0, False)
        sheet = wb.Worksheets("Design Regulator")

        # Known input cells (observed in the template)
        sheet.Range("E6").Value2 = inp.vin_min_v
        sheet.Range("E7").Value2 = inp.vin_nom_v
        sheet.Range("E8").Value2 = inp.vin_max_v
        sheet.Range("E9").Value2 = inp.vout_v
        sheet.Range("E10").Value2 = inp.iout_a
        sheet.Range("E11").Value2 = fsw_khz

        # Force a full recalc so dependent sheets update
        excel.CalculateFullRebuild()

        # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        wb.SaveAs(str(out_xlsm), 52)

        if out_xlsx is not None:
            # 51 = xlOpenXMLWorkbook (.xlsx)
            wb.SaveAs(str(out_xlsx), 51)

    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def main() -> int:
    p = argparse.ArgumentParser(
        description="Fill TI LM5148 quickstart calculator (.xlsm) from webapp JSON, using real Excel via COM."
    )
    p.add_argument("--json", required=True, help="Path to lm5148_design.json exported by the web app")
    p.add_argument(
        "--template",
        default=str(Path(__file__).resolve().parent.parent / "training" / "LM5148_LM25148_quickstart_calculator_A4.xlsm"),
        help="Path to TI .xlsm template",
    )
    p.add_argument("--out-xlsm", default=str(Path(__file__).resolve().parent / "LM5148_quickstart_filled_excel.xlsm"))
    p.add_argument("--out-xlsx", default=str(Path(__file__).resolve().parent / "LM5148_quickstart_filled_excel.xlsx"))
    p.add_argument("--no-xlsx", action="store_true", help="Do not write the .xlsx copy")
    p.add_argument("--visible", action="store_true", help="Show the Excel UI while running")

    args = p.parse_args()

    json_path = Path(args.json)
    template_path = Path(args.template)
    out_xlsm = Path(args.out_xlsm)
    out_xlsx = None if args.no_xlsx else Path(args.out_xlsx)

    fill_quickstart_excel(
        json_path=json_path,
        template_path=template_path,
        out_xlsm=out_xlsm,
        out_xlsx=out_xlsx,
        visible=bool(args.visible),
    )
    print(f"Wrote: {out_xlsm}")
    if out_xlsx is not None:
        print(f"Wrote: {out_xlsx}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
