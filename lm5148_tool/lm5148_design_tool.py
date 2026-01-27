from __future__ import annotations

import argparse
import math
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
import xlsxwriter


VREF_DEFAULT_V = 0.8


@dataclass(frozen=True)
class DesignInputs:
    vin_nom_v: float = 12.0
    vin_max_v: float = 18.0
    vout_v: float = 5.0
    iout_a: float = 8.0
    fsw_hz: float = 2.1e6

    # Inductor ripple current target as fraction of IOUT (datasheet uses ~30%).
    ripple_frac: float = 0.30

    # If provided, use this inductance for peak-current / ripple checks.
    l_used_h: Optional[float] = 0.56e-6

    # Output transient spec for load-off (Vout overshoot allowed).
    vout_overshoot_v: float = 0.075

    # Capacitor ESR assumptions
    rout_esr_ohm: float = 1e-3
    rin_esr_ohm: float = 2e-3

    # Input ripple spec
    vin_ripple_pp_v: float = 0.120

    # Current limit / timing assumptions
    vcs_th_v: float = 0.060
    il_pk_margin: float = 1.25
    t_delay_isns_s: float = 45e-9

    # Feedback design
    vref_v: float = VREF_DEFAULT_V
    rfb_bottom_ohm: float = 10_000.0

    # Compensation starting point (datasheet pages 38-39)
    f_c_hz: float = 60_000.0
    rcomp_ohm: float = 10_000.0
    f_esr_zero_hz: float = 500_000.0
    cbw_f: float = 0.8e-12

    # Paths
    pdf_path: Optional[str] = None


@dataclass(frozen=True)
class DesignResults:
    delta_il_nom_a: float
    l_required_h: float

    delta_il_vin_max_a: float
    il_peak_vin_max_a: float

    rsense_ohm: float
    il_peak_short_a: float

    cout_load_off_f: float
    vout_ripple_pp_v: float
    ioutcap_rms_a: float

    duty_nom: float
    cin_rms_a: float
    cin_required_f: float

    rt_ohm: float
    rfb_top_ohm: float

    ccomp_f: float
    chf_f: float


def _clamp(x: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, x))


def eq31_l_required(vin_nom_v: float, vout_v: float, fsw_hz: float, delta_il_a: float) -> float:
    # Reconstructed from datasheet example (page 36, Eq.31):
    # L = Vout*(Vin-Vout)/(Vin*Fsw*ΔIL)
    return (vout_v * (vin_nom_v - vout_v)) / (vin_nom_v * fsw_hz * delta_il_a)


def inductor_ripple(vin_v: float, vout_v: float, fsw_hz: float, l_h: float) -> float:
    # ΔIL = Vout*(Vin-Vout)/(Vin*L*Fsw)
    return (vout_v * (vin_v - vout_v)) / (vin_v * l_h * fsw_hz)


def eq32_il_peak(vin_max_v: float, vout_v: float, fsw_hz: float, l_h: float, iout_a: float) -> tuple[float, float]:
    delta_il = inductor_ripple(vin_max_v, vout_v, fsw_hz, l_h)
    il_pk = iout_a + (delta_il / 2.0)
    return delta_il, il_pk


def eq34_rsense(vcs_th_v: float, il_pk_a: float, margin: float) -> float:
    # Rsense = Vcs_th / (margin * IL_pk)
    return vcs_th_v / (margin * il_pk_a)


def eq35_il_peak_short(vin_max_v: float, t_delay_s: float, vcs_th_v: float, rsense_ohm: float, l_h: float) -> float:
    # IL_pk(sc) ≈ Vcs_th/Rsense + Vin_max * t_delay / L
    return (vcs_th_v / rsense_ohm) + (vin_max_v * t_delay_s / l_h)


def eq36_cout_load_off(l_h: float, iout_a: float, vout_v: float, overshoot_v: float) -> float:
    # Energy transfer approximation (matches datasheet numeric example):
    # 1/2 L I^2 = 1/2 C * ((Vout+ΔV)^2 - Vout^2) = 1/2 C * (2 Vout ΔV + ΔV^2)
    denom = (2.0 * vout_v * overshoot_v) + (overshoot_v**2)
    return (l_h * (iout_a**2)) / denom


def eq37_vout_ripple_pp(delta_il_a: float, fsw_hz: float, cout_eff_f: float, rout_esr_ohm: float) -> float:
    # Vripple_pp ≈ RSS( ΔIL/(8 Fsw C) , ΔIL * ESR )
    v_c = delta_il_a / (8.0 * fsw_hz * cout_eff_f)
    v_esr = delta_il_a * rout_esr_ohm
    return math.sqrt(v_c * v_c + v_esr * v_esr)


def eq38_ioutcap_rms(delta_il_a: float) -> float:
    return delta_il_a / math.sqrt(12.0)


def eq39_cin_rms(iout_a: float, duty: float) -> float:
    return iout_a * math.sqrt(duty * (1.0 - duty))


def eq40_cin_required(iout_a: float, fsw_hz: float, duty: float, dv_in_pp_v: float, rin_esr_ohm: float) -> float:
    # Reconstructed to match datasheet behavior: compute Cin required so that
    # total input ripple is within dv_in_pp_v, combining capacitive + ESR ripple in RSS.
    # For a triangular capacitor current, the charge/discharge contribution is approximated as:
    # ΔVin_cap ≈ Iout*D*(1-D) / (Fsw * Cin)
    i_factor = iout_a * duty * (1.0 - duty)
    i_cin_rms = eq39_cin_rms(iout_a, duty)
    dv_esr = i_cin_rms * rin_esr_ohm

    # Guard: if ESR ripple alone exceeds spec, capacitance can't fix it.
    if dv_esr >= dv_in_pp_v:
        return float("inf")

    dv_cap_allow = math.sqrt(max(dv_in_pp_v**2 - dv_esr**2, 0.0))
    if dv_cap_allow <= 0:
        return float("inf")

    return i_factor / (fsw_hz * dv_cap_allow)


def eq41_rt_ohm_from_fsw(fsw_hz: float) -> float:
    # Datasheet Eq.41 reconstructed from constants shown:
    # Fsw(kHz) = 1e6 / (45*Rt(kΩ) + 53)
    fsw_khz = fsw_hz / 1_000.0
    rt_kohm = (1_000_000.0 / fsw_khz - 53.0) / 45.0
    return rt_kohm * 1_000.0


def eq42_feedback_top(vout_v: float, vref_v: float, r_bottom_ohm: float) -> float:
    # Standard divider: Vout = Vref * (1 + Rtop/Rbottom)
    if vout_v <= vref_v:
        return 0.0
    return r_bottom_ohm * (vout_v / vref_v - 1.0)


def eq44_ccomp(f_c_hz: float, rcomp_ohm: float) -> float:
    # Datasheet (page 38, Eq.44): place compensation zero at f_c/10.
    # Ccomp = 10 / (2π f_c Rcomp)
    return 10.0 / (2.0 * math.pi * f_c_hz * rcomp_ohm)


def eq45_chf(f_esr_zero_hz: float, rcomp_ohm: float, cbw_f: float) -> float:
    # Datasheet (page 39, Eq.45): CHF = 1/(2π f_ESR Rcomp) - Cbw
    return (1.0 / (2.0 * math.pi * f_esr_zero_hz * rcomp_ohm)) - cbw_f


def run_design(inp: DesignInputs) -> DesignResults:
    duty_nom = _clamp(inp.vout_v / inp.vin_nom_v, 0.0, 0.95)

    delta_il_nom = inp.ripple_frac * inp.iout_a
    l_req = eq31_l_required(inp.vin_nom_v, inp.vout_v, inp.fsw_hz, delta_il_nom)

    l_used = inp.l_used_h if inp.l_used_h is not None else l_req
    delta_il_vin_max, il_pk_vin_max = eq32_il_peak(inp.vin_max_v, inp.vout_v, inp.fsw_hz, l_used, inp.iout_a)

    rsense = eq34_rsense(inp.vcs_th_v, il_pk_vin_max, inp.il_pk_margin)
    il_pk_short = eq35_il_peak_short(inp.vin_max_v, inp.t_delay_isns_s, inp.vcs_th_v, rsense, l_used)

    cout_load_off = eq36_cout_load_off(l_used, inp.iout_a, inp.vout_v, inp.vout_overshoot_v)

    # If user didn't provide an effective output capacitance, use Cout from eq36 as a baseline.
    cout_eff = cout_load_off
    vout_ripple = eq37_vout_ripple_pp(delta_il_nom, inp.fsw_hz, cout_eff, inp.rout_esr_ohm)
    ioutcap_rms = eq38_ioutcap_rms(delta_il_nom)

    cin_rms = eq39_cin_rms(inp.iout_a, 0.5)
    cin_req = eq40_cin_required(inp.iout_a, inp.fsw_hz, 0.5, inp.vin_ripple_pp_v, inp.rin_esr_ohm)

    rt_ohm = eq41_rt_ohm_from_fsw(inp.fsw_hz)
    rfb_top = eq42_feedback_top(inp.vout_v, inp.vref_v, inp.rfb_bottom_ohm)

    ccomp = eq44_ccomp(inp.f_c_hz, inp.rcomp_ohm)
    chf = eq45_chf(inp.f_esr_zero_hz, inp.rcomp_ohm, inp.cbw_f)

    return DesignResults(
        delta_il_nom_a=delta_il_nom,
        l_required_h=l_req,
        delta_il_vin_max_a=delta_il_vin_max,
        il_peak_vin_max_a=il_pk_vin_max,
        rsense_ohm=rsense,
        il_peak_short_a=il_pk_short,
        cout_load_off_f=cout_load_off,
        vout_ripple_pp_v=vout_ripple,
        ioutcap_rms_a=ioutcap_rms,
        duty_nom=duty_nom,
        cin_rms_a=cin_rms,
        cin_required_f=cin_req,
        rt_ohm=rt_ohm,
        rfb_top_ohm=rfb_top,
        ccomp_f=ccomp,
        chf_f=chf,
    )


def extract_equation_images(
    pdf_path: Path,
    out_dir: Path,
    equation_numbers: list[int],
    pages_1based: tuple[int, int] = (36, 39),
    dpi: int = 220,
) -> dict[int, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    eq_to_path: dict[int, Path] = {}

    with fitz.open(pdf_path) as doc:
        for eq in equation_numbers:
            token = f"({eq})"
            found = False
            for pno in range(pages_1based[0] - 1, pages_1based[1]):
                page = doc.load_page(pno)
                rects = page.search_for(token)
                if not rects:
                    continue

                # Use the first match.
                r = rects[0]
                # Crop a region above the equation number that typically contains the full equation.
                clip = fitz.Rect(
                    0,
                    max(0, r.y0 - 180),
                    page.rect.width,
                    min(page.rect.height, r.y1 + 30),
                )
                pix = page.get_pixmap(clip=clip, dpi=dpi)
                out_path = out_dir / f"eq_{eq}_p{pno+1}.png"
                pix.save(out_path)
                eq_to_path[eq] = out_path
                found = True
                break

            if not found:
                # Skip silently; some equations (e.g., 42) can be graphical and harder to search.
                continue

    return eq_to_path


def export_to_excel(
    inp: DesignInputs,
    res: DesignResults,
    out_xlsx: Path,
    equation_images: dict[int, Path],
) -> None:
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)

    workbook = xlsxwriter.Workbook(str(out_xlsx))
    try:
        ws_in = workbook.add_worksheet("Inputs")
        ws_out = workbook.add_worksheet("Results")
        ws_eq = workbook.add_worksheet("Equations")

        header_fmt = workbook.add_format({"bold": True, "bg_color": "#E6E6E6"})
        num_fmt = workbook.add_format({"num_format": "0.0000"})
        sci_fmt = workbook.add_format({"num_format": "0.00E+00"})

        # Inputs
        ws_in.write_row(0, 0, ["Parameter", "Value", "Units"], header_fmt)
        r = 1
        for k, v in asdict(inp).items():
            if k == "pdf_path":
                continue
            ws_in.write(r, 0, k)
            ws_in.write(r, 1, v)
            ws_in.write(r, 2, "")
            r += 1
        ws_in.set_column(0, 0, 26)
        ws_in.set_column(1, 1, 18)
        ws_in.set_column(2, 2, 10)

        # Results
        ws_out.write_row(0, 0, ["Item", "Value", "Units", "Notes"], header_fmt)
        results_rows = [
            ["ΔIL @ Vin_nom (target)", res.delta_il_nom_a, "A", "Eq.31 uses this ripple target"],
            ["L required", res.l_required_h, "H", "Eq.31"],
            ["ΔIL @ Vin_max (with L_used)", res.delta_il_vin_max_a, "A", "Eq.32"],
            ["IL peak @ Vin_max", res.il_peak_vin_max_a, "A", "Eq.32"],
            ["Rsense", res.rsense_ohm, "Ω", "Eq.34"],
            ["IL peak short", res.il_peak_short_a, "A", "Eq.35"],
            ["Cout load-off (min)", res.cout_load_off_f, "F", "Eq.36"],
            ["Vout ripple pp (est)", res.vout_ripple_pp_v, "Vpp", "Eq.37"],
            ["Ioutcap RMS", res.ioutcap_rms_a, "A", "Eq.38"],
            ["Duty @ Vin_nom", res.duty_nom, "", "Vout/Vin_nom"],
            ["Cin RMS (D=0.5)", res.cin_rms_a, "A", "Eq.39"],
            ["Cin required (D=0.5)", res.cin_required_f, "F", "Eq.40"],
            ["RT", res.rt_ohm, "Ω", "Eq.41"],
            ["Rfb top (given Rbottom)", res.rfb_top_ohm, "Ω", "Eq.42 (standard divider)"],
            ["RCOMP (given)", inp.rcomp_ohm, "Ω", "Starting value from datasheet procedure"],
            ["CCOMP", res.ccomp_f, "F", "Eq.44: zero at fC/10"],
            ["CHF", res.chf_f, "F", "Eq.45: pole at ESR zero (minus Cbw)"],
        ]
        for i, row in enumerate(results_rows, start=1):
            ws_out.write(i, 0, row[0])
            # Use scientific notation for very small/large values.
            value = row[1]
            fmt = sci_fmt if (isinstance(value, (int, float)) and (abs(value) < 1e-3 or abs(value) >= 1e4)) else num_fmt
            ws_out.write(i, 1, value, fmt)
            ws_out.write(i, 2, row[2])
            ws_out.write(i, 3, row[3])
        ws_out.set_column(0, 0, 28)
        ws_out.set_column(1, 1, 18)
        ws_out.set_column(2, 2, 10)
        ws_out.set_column(3, 3, 45)

        # Equations
        ws_eq.write_row(0, 0, ["Equation", "Image file", "Image"], header_fmt)
        row = 1
        for eq in sorted(equation_images.keys()):
            ws_eq.write(row, 0, f"({eq})")
            ws_eq.write(row, 1, str(equation_images[eq]))
            # Insert image in column C; scale down to fit.
            ws_eq.insert_image(row, 2, str(equation_images[eq]), {"x_scale": 0.6, "y_scale": 0.6})
            row += 8
        ws_eq.set_column(0, 0, 10)
        ws_eq.set_column(1, 1, 70)
        ws_eq.set_column(2, 2, 60)
    finally:
        workbook.close()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="LM5148 design helper based on datasheet pages 36-39; exports an Excel summary."
    )

    parser.add_argument("--vin-nom", type=float, default=DesignInputs.vin_nom_v)
    parser.add_argument("--vin-max", type=float, default=DesignInputs.vin_max_v)
    parser.add_argument("--vout", type=float, default=DesignInputs.vout_v)
    parser.add_argument("--iout", type=float, default=DesignInputs.iout_a)
    parser.add_argument("--fsw", type=float, default=DesignInputs.fsw_hz, help="Switching frequency in Hz")
    parser.add_argument("--ripple-frac", type=float, default=DesignInputs.ripple_frac)
    parser.add_argument("--l-used", type=float, default=DesignInputs.l_used_h, help="Inductance in H")
    parser.add_argument("--vout-overshoot", type=float, default=DesignInputs.vout_overshoot_v)
    parser.add_argument("--vin-ripple", type=float, default=DesignInputs.vin_ripple_pp_v)
    parser.add_argument("--rout-esr", type=float, default=DesignInputs.rout_esr_ohm)
    parser.add_argument("--rin-esr", type=float, default=DesignInputs.rin_esr_ohm)
    parser.add_argument("--vcs-th", type=float, default=DesignInputs.vcs_th_v)
    parser.add_argument("--il-pk-margin", type=float, default=DesignInputs.il_pk_margin)
    parser.add_argument("--t-delay", type=float, default=DesignInputs.t_delay_isns_s)
    parser.add_argument("--vref", type=float, default=DesignInputs.vref_v)
    parser.add_argument("--rfb-bottom", type=float, default=DesignInputs.rfb_bottom_ohm)

    parser.add_argument("--fc", type=float, default=DesignInputs.f_c_hz, help="Loop crossover frequency in Hz")
    parser.add_argument("--rcomp", type=float, default=DesignInputs.rcomp_ohm, help="Compensation resistor in ohms")
    parser.add_argument(
        "--fesr",
        type=float,
        default=DesignInputs.f_esr_zero_hz,
        help="ESR-zero frequency in Hz (used for CHF calculation)",
    )
    parser.add_argument(
        "--cbw",
        type=float,
        default=DesignInputs.cbw_f,
        help="Error amplifier bandwidth limiting capacitor Cbw in farads",
    )

    parser.add_argument(
        "--pdf",
        type=str,
        default=str(
            Path(
                r"m:\Amit Medina\Schematic and Layout Design Course\Schematic-and-Layout-Design-Course-repo\HW Training - Schematic to Layout Design Course\lm5148.pdf"
            )
        ),
        help="Path to lm5148.pdf (used to embed equation images)",
    )
    parser.add_argument(
        "--out",
        type=str,
        default=str(Path.cwd() / "lm5148_design_export.xlsx"),
        help="Output .xlsx path",
    )

    args = parser.parse_args()

    inp = DesignInputs(
        vin_nom_v=args.vin_nom,
        vin_max_v=args.vin_max,
        vout_v=args.vout,
        iout_a=args.iout,
        fsw_hz=args.fsw,
        ripple_frac=args.ripple_frac,
        l_used_h=args.l_used,
        vout_overshoot_v=args.vout_overshoot,
        rout_esr_ohm=args.rout_esr,
        rin_esr_ohm=args.rin_esr,
        vin_ripple_pp_v=args.vin_ripple,
        vcs_th_v=args.vcs_th,
        il_pk_margin=args.il_pk_margin,
        t_delay_isns_s=args.t_delay,
        vref_v=args.vref,
        rfb_bottom_ohm=args.rfb_bottom,
        f_c_hz=args.fc,
        rcomp_ohm=args.rcomp,
        f_esr_zero_hz=args.fesr,
        cbw_f=args.cbw,
        pdf_path=args.pdf,
    )

    res = run_design(inp)

    pdf_path = Path(args.pdf)
    images_dir = Path.cwd() / "lm5148_equations_images_v3"
    eq_images = {}
    if pdf_path.exists():
        eq_images = extract_equation_images(
            pdf_path,
            images_dir,
            equation_numbers=[31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 43, 44, 45],
        )

    export_to_excel(inp, res, Path(args.out), eq_images)
    print(f"Wrote Excel: {args.out}")
    if eq_images:
        print(f"Wrote {len(eq_images)} equation images to: {images_dir}")
    else:
        print("No equation images embedded (PDF missing or equations not found).")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
