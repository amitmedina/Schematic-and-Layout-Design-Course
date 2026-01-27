from __future__ import annotations

import io
from dataclasses import asdict
from typing import Any

import xlsxwriter


def build_results_xlsx_bytes(*, inputs: dict[str, Any], results: dict[str, Any]) -> bytes:
    """Create a simple, standalone .xlsx report as bytes.

    This is intentionally lightweight so it can be used from Streamlit (download button)
    without touching Excel/COM.
    """

    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True})

    ws = wb.add_worksheet("LM5148 Results")

    header = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    key_fmt = wb.add_format({"bold": True})
    num_fmt = wb.add_format({"num_format": "0.000000"})

    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 22)

    row = 0
    ws.write(row, 0, "Inputs", header)
    row += 1

    for k in sorted(inputs.keys()):
        v = inputs[k]
        ws.write(row, 0, k, key_fmt)
        if isinstance(v, (int, float)):
            ws.write_number(row, 1, float(v), num_fmt)
        else:
            ws.write(row, 1, str(v))
        row += 1

    row += 1
    ws.write(row, 0, "Results", header)
    row += 1

    for k in sorted(results.keys()):
        v = results[k]
        ws.write(row, 0, k, key_fmt)
        if isinstance(v, (int, float)):
            ws.write_number(row, 1, float(v), num_fmt)
        else:
            ws.write(row, 1, str(v))
        row += 1

    wb.close()
    return bio.getvalue()
