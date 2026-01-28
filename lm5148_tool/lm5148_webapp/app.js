const STORAGE_KEY = 'lm5148_webapp_inputs_v1';

const DEFAULTS = {
  vinMin: 10,
  vinNom: 12,
  vinMax: 18,
  vout: 5,
  iout: 8,
  fsw: 2.1e6,

  rippleFrac: 0.30,
  lUsed: 0.56e-6,
  lockLUsedToLreq: true,

  // Eq33 helper
  // UI: mΩ
  rsForEq33: 5,

  // Current sense
  vcsTh: 0.060,
  ilPkMargin: 1.25,
  tDelay: 45e-9,

  // Output caps
  voutOvershoot: 0.075,
  coutEff: 44e-6,
  routEsr: 1e-3,

  // Input caps
  duty: 0.5,
  vinRippleSpec: 0.120,
  rinEsr: 2e-3,

  // Feedback
  vref: 0.8,
  // Schematic/manual choice (not used by Eq.42 math output)
  rfbTopManual: 48_700,
  rfbBot: 10_000,

  // Compensation
  fc: 60_000,
  gm: 1200e-6,
  fesr: 500_000,
  cbw: 0.8e-12,

  // Quickstart-style schematic components (display + interactive diagram)
  // Note: CIN and COUT are calculated from Eq.40 and Eq.36 respectively
  uvloR1: 200_000,
  uvloR2: 9_530,
  // UVLO calculator (Eq.2–3 style)
  uvloVinOn: 10,
  uvloVinOff: 9,
  uvloIhys: 10e-6,
  uvloVen: 1.0,
  uvloLock: false,
  cvinSmall: 22e-9,
  cvcc: 2.2e-6,
  cboot: 0.1e-6,
  rcnfg: 41_200,
  cvdda: 0.1e-6,
};

function uvlo_ruv1(vinOn, vinOff, ihys) {
  if (!Number.isFinite(vinOn) || !Number.isFinite(vinOff) || !Number.isFinite(ihys)) return NaN;
  if (ihys <= 0) return NaN;
  if (vinOn <= vinOff) return NaN;
  return (vinOn - vinOff) / ihys;
}

function uvlo_ruv2(ruv1, ven, vinOn) {
  if (!Number.isFinite(ruv1) || !Number.isFinite(ven) || !Number.isFinite(vinOn)) return NaN;
  if (ruv1 <= 0) return NaN;
  if (vinOn <= ven) return NaN;
  return ruv1 * (ven / (vinOn - ven));
}

let lastResults = null;

function $(id) {
  return document.getElementById(id);
}

function setStatus(msg) {
  const el = $('status');
  if (!el) return;
  el.textContent = msg || '';
}

function safeNumber(x) {
  const v = Number(x);
  return Number.isFinite(v) ? v : NaN;
}

function fmt(v, digits = 6) {
  if (!Number.isFinite(v)) return '—';
  const abs = Math.abs(v);
  if (abs === 0) return '0';
  if (abs < 1e-3 || abs >= 1e4) return v.toExponential(3);
  return v.toFixed(digits).replace(/\.0+$/, '').replace(/(\.[0-9]*?)0+$/, '$1');
}

function fmtEng(value, unit = '') {
  if (!Number.isFinite(value)) return '—';
  if (value === 0) return unit ? `0 ${unit}` : '0';

  const abs = Math.abs(value);
  const prefixes = [
    { exp: -12, p: 'p' },
    { exp: -9, p: 'n' },
    { exp: -6, p: 'µ' },
    { exp: -3, p: 'm' },
    { exp: 0, p: '' },
    { exp: 3, p: 'k' },
    { exp: 6, p: 'M' },
    { exp: 9, p: 'G' },
    { exp: 12, p: 'T' },
  ];

  const exp = Math.floor(Math.log10(abs) / 3) * 3;
  const clamped = Math.max(-12, Math.min(12, exp));
  const prefix = prefixes.find(x => x.exp === clamped) ?? prefixes[4];

  const scaled = value / (10 ** prefix.exp);
  const scaledAbs = Math.abs(scaled);

  // Prefer 3–4 significant digits.
  const digits = scaledAbs < 10 ? 3 : (scaledAbs < 100 ? 2 : 1);
  const num = Number(scaled.toFixed(digits));
  const suffix = unit ? `${prefix.p}${unit}` : '';
  return suffix ? `${num} ${suffix}` : String(num);
}

function setText(id, text) {
  const el = $(id);
  if (!el) return;
  el.textContent = text;
}

function setEng(id, value, unit) {
  setText(id, fmtEng(value, unit));
}

function hookSchematicHotspots() {
  const svg = $('schematicSvg');
  if (!svg) return;
  const hotspots = svg.querySelectorAll('[data-target]');
  hotspots.forEach((el) => {
    el.addEventListener('click', () => {
      const targetId = el.getAttribute('data-target');
      if (!targetId) return;
      const target = $(targetId);
      if (!target) return;
      target.scrollIntoView({ behavior: 'smooth', block: 'center' });
      try {
        target.focus({ preventScroll: true });
      } catch {
        target.focus();
      }
    });
  });
}

function preferOptionalImages() {
  const imgs = document.querySelectorAll('img[data-prefer-src]');
  imgs.forEach((img) => {
    const preferred = img.getAttribute('data-prefer-src');
    if (!preferred) return;
    const probe = new Image();
    probe.onload = () => {
      img.src = preferred;
    };
    // If it fails, keep the existing SVG.
    probe.onerror = () => {};
    probe.src = preferred;
  });
}

function updateSchematic(st, results, rsense_Ohm) {
  // System
  setText('sch_vin', `${fmtShort(st.vinMin)}V – ${fmtShort(st.vinMax)}V`);
  setText('sch_vout', `${fmtShort(st.vout)}V`);
  setText('sch_iout', `${fmtShort(st.iout)}A`);

  // Inductor: show the L used value (calculated from Eq.31 when locked)
  setText('sch_l', fmtEng(st.lUsed, 'H'));

  // Output capacitor: show calculated Cout min from Eq.36
  setText('sch_cout', fmtEng(results?.coutMin_F, 'F'));
  setText('sch_cout_eff', `eff: ${fmtEng(st.coutEff, 'F')}`);

  // Input capacitor: show calculated CIN required from Eq.40
  const cinReq = results?.cinReq_F;
  setText('sch_cin', (cinReq === Infinity) ? '∞' : fmtEng(cinReq, 'F'));

  // Quickstart support parts (manual values)
  setText('sch_cvin', fmtEng(st.cvinSmall, 'F'));
  setText('sch_cvcc', fmtEng(st.cvcc, 'F'));
  setText('sch_cboot', fmtEng(st.cboot, 'F'));
  setText('sch_rcnfg', fmtEng(st.rcnfg, 'Ω'));
  setText('sch_cvdda', fmtEng(st.cvdda, 'F'));
  setText('sch_ruv1', fmtEng(st.uvloR1, 'Ω'));
  setText('sch_ruv2', fmtEng(st.uvloR2, 'Ω'));

  // Calculated values
  setText('sch_rs', fmtEng(rsense_Ohm, 'Ω'));
  setText('sch_rt', fmtEng(results?.rt_Ohm, 'Ω'));

  // RFB shown as a manual schematic value (matches quickstart-style schematic)
  setText('sch_rtop', fmtEng(results?.rfbTop_Ohm, 'Ω'));
  setText('sch_rbot', `RFB2: ${fmtEng(st.rfbBot, 'Ω')}`);

  setText('sch_rcomp', fmtEng(results?.rcomp_Ohm, 'Ω'));
  setText('sch_ccomp', fmtEng(results?.ccomp_F, 'F'));
  setText('sch_chf', fmtEng(results?.chf_F, 'F'));
}

function fmtShort(v) {
  return fmt(v, 6);
}

function fmtHz(v) {
  // For readability in the substitution line, prefer plain Hz when it's not huge.
  if (!Number.isFinite(v)) return '—';
  if (v >= 1e6) return fmtEng(v, 'Hz');
  return fmt(v, 0);
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULTS };
    const obj = JSON.parse(raw);
    const st = { ...DEFAULTS, ...obj };

    // Backward-compatible migration:
    // Older versions stored rsForEq33 in Ω (e.g., 0.005). New UI expects mΩ (e.g., 5).
    if (typeof st.rsForEq33 === 'number' && Number.isFinite(st.rsForEq33) && st.rsForEq33 > 0 && st.rsForEq33 < 0.2) {
      st.rsForEq33 = st.rsForEq33 * 1000;
    }

    return st;
  } catch {
    return { ...DEFAULTS };
  }
}

function saveState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function readInputs() {
  const state = {};
  for (const key of Object.keys(DEFAULTS)) {
    const el = $(key);
    if (!el) continue;
    if (el.type === 'checkbox') {
      state[key] = !!el.checked;
    } else {
      state[key] = safeNumber(el.value);
    }
  }
  return state;
}

function writeInputs(state) {
  for (const key of Object.keys(DEFAULTS)) {
    const el = $(key);
    if (!el) continue;
    if (el.type === 'checkbox') {
      el.checked = !!state[key];
    } else {
      el.value = state[key];
    }
  }
}

// Equations
function eq31_L(vinNom, vout, fsw, deltaIl) {
  return (vout * (vinNom - vout)) / (vinNom * fsw * deltaIl);
}

function rippleDeltaIl(iout, rippleFrac) {
  return iout * rippleFrac;
}

function deltaIl(vin, vout, fsw, L) {
  return (vout * (vin - vout)) / (vin * L * fsw);
}

function eq32_ilPk(iout, deltaIlMax) {
  return iout + deltaIlMax / 2;
}

function eq33_lo_sc(vout, rs_mOhm, fsw) {
  // Datasheet Eq. 33: L_O(sc) = (VOUT * RS) / ((24 mV) * FSW)
  // This matches the datasheet numeric example (~0.5 µH for 5V, 5mΩ, 2.1MHz).
  // UI expects RS in mΩ; convert to Ω here.
  const rs_ohm = rs_mOhm * 1e-3;
  const slope_v = 24e-3;
  return (vout * rs_ohm) / (slope_v * fsw);
}

function eq34_rsense(vcsTh, margin, ilPk) {
  return vcsTh / (margin * ilPk);
}

function eq35_ilPkSc(vcsTh, rs, vinMax, tDelay, L) {
  return (vcsTh / rs) + (vinMax * tDelay / L);
}

function eq36_coutMin(L, iout, vout, dV) {
  const denom = (vout + dV) ** 2 - vout ** 2;
  return (L * iout ** 2) / denom;
}

function eq37_voutRipple(deltaIlNom, fsw, coutEff, rEsr) {
  const vCap = deltaIlNom / (8 * fsw * coutEff);
  const vEsr = deltaIlNom * rEsr;
  return Math.sqrt(vCap ** 2 + vEsr ** 2);
}

function eq38_icoRms(deltaIlNom) {
  return deltaIlNom / Math.sqrt(12);
}

function eq39_icinRms(iout, D) {
  return iout * Math.sqrt(D * (1 - D));
}

function eq40_cinRequired(iout, D, fsw, dvSpec, rinEsr) {
  const icinRms = eq39_icinRms(iout, D);
  const dvEsr = icinRms * rinEsr;
  if (dvEsr >= dvSpec) return { cin: Infinity, note: 'ESR ripple alone exceeds the ripple spec.' };
  const dvCapAllow = Math.sqrt(Math.max(dvSpec ** 2 - dvEsr ** 2, 0));
  if (dvCapAllow <= 0) return { cin: Infinity, note: 'No headroom left for capacitive ripple.' };
  const iFactor = iout * D * (1 - D);
  const cin = iFactor / (fsw * dvCapAllow);
  return { cin, note: '' };
}

function eq41_rtOhm(fswHz) {
  const fswKHz = fswHz / 1000;
  const rtKOhm = (1_000_000 / fswKHz - 53) / 45;
  return rtKOhm * 1000;
}

function eq42_rtop(vout, vref, rbot) {
  if (vout <= vref) return 0;
  return rbot * (vout / vref - 1);
}

function eq43_rcomp(vout, rs, gm, fc, coutEff, vref) {
  return (vout * rs * gm) / (2 * Math.PI * fc * coutEff * vref);
}

function eq44_ccomp(fc, rcomp) {
  return 10 / (2 * Math.PI * fc * rcomp);
}

function eq45_chf(fesr, rcomp, cbw) {
  return 1 / (2 * Math.PI * fesr * rcomp) - cbw;
}

function recalc() {
  const st = readInputs();

  // UVLO calculator (Eq.2–3) — optionally drives schematic RUV1/RUV2
  const uvloR1Calc = uvlo_ruv1(st.uvloVinOn, st.uvloVinOff, st.uvloIhys);
  const uvloR2Calc = uvlo_ruv2(uvloR1Calc, st.uvloVen, st.uvloVinOn);

  setEng('uvloCalcR1', uvloR1Calc, 'Ω');
  setEng('uvloCalcR2', uvloR2Calc, 'Ω');

  setText(
    'uvlo_eq2_sub',
    `RUV1 = (VIN(on)−VIN(off))/IHYS = (${fmtShort(st.uvloVinOn)}−${fmtShort(st.uvloVinOff)})/${fmtEng(st.uvloIhys, 'A')} = ${fmtEng(uvloR1Calc, 'Ω')}`
  );
  setText(
    'uvlo_eq3_sub',
    `RUV2 = RUV1·VEN/(VIN(on)−VEN) = ${fmtEng(uvloR1Calc, 'Ω')}·${fmtShort(st.uvloVen)}/(${fmtShort(st.uvloVinOn)}−${fmtShort(st.uvloVen)}) = ${fmtEng(uvloR2Calc, 'Ω')}`
  );

  const uvloR1El = $('uvloR1');
  const uvloR2El = $('uvloR2');
  if (st.uvloLock && Number.isFinite(uvloR1Calc) && Number.isFinite(uvloR2Calc)) {
    st.uvloR1 = uvloR1Calc;
    st.uvloR2 = uvloR2Calc;
    if (uvloR1El) uvloR1El.value = String(uvloR1Calc);
    if (uvloR2El) uvloR2El.value = String(uvloR2Calc);
  }

  // Derived values
  const dIlNom = rippleDeltaIl(st.iout, st.rippleFrac);
  $('deltaIlNom').value = Number.isFinite(dIlNom) ? dIlNom : '';

  // Substitution line for ΔIL (used by Eq.31)
  setText(
    'deltaIl_sub',
    `ΔIL = rippleFrac·IOUT = ${fmtShort(st.rippleFrac)}·${fmtShort(st.iout)} A = ${fmtEng(dIlNom, 'A')}`
  );

  const Lreq = eq31_L(st.vinNom, st.vout, st.fsw, dIlNom);
  setEng('lReq', Lreq, 'H');

  setText(
    'eq31_sub',
    `L = VOUT·(VIN_nom−VOUT)/(VIN_nom·FSW·ΔIL) = ${fmtShort(st.vout)}·(${fmtShort(st.vinNom)}−${fmtShort(st.vout)})/(${fmtShort(st.vinNom)}·${fmtHz(st.fsw)}·${fmtShort(dIlNom)}) = ${fmtEng(Lreq, 'H')}`
  );

  // By default, drive "L used" from Eq.31 result.
  const lUsedEl = $('lUsed');
  const lockEl = $('lockLUsedToLreq');
  if (st.lockLUsedToLreq) {
    st.lUsed = Lreq;
    if (lUsedEl) {
      lUsedEl.value = String(Lreq);
      lUsedEl.disabled = true;
    }
  } else {
    if (lUsedEl) lUsedEl.disabled = false;
  }

  const dIlMax = deltaIl(st.vinMax, st.vout, st.fsw, st.lUsed);
  setEng('deltaIlMax', dIlMax, 'A');

  const ilPkMax = eq32_ilPk(st.iout, dIlMax);
  setEng('ilPkMax', ilPkMax, 'A');

  setText(
    'eq32_sub',
    `ΔIL = VOUT·(VIN_max−VOUT)/(VIN_max·L·FSW) = ${fmtShort(st.vout)}·(${fmtShort(st.vinMax)}−${fmtShort(st.vout)})/(${fmtShort(st.vinMax)}·${fmtEng(st.lUsed, 'H')}·${fmtHz(st.fsw)}) = ${fmtEng(dIlMax, 'A')};  IL,pk = IOUT + ΔIL/2 = ${fmtShort(st.iout)} + ${fmtShort(dIlMax)}/2 = ${fmtEng(ilPkMax, 'A')}`
  );

  const loSc = eq33_lo_sc(st.vout, st.rsForEq33, st.fsw);
  setEng('rosc', loSc, 'H');

  setText(
    'eq33_sub',
    `LO(sc) = (VOUT·RS)/((24 mV)·FSW) = (${fmtShort(st.vout)}·${fmtShort(st.rsForEq33)} mΩ)/(24 mV·${fmtHz(st.fsw)}) = ${fmtEng(loSc, 'H')}`
  );

  const rsense = eq34_rsense(st.vcsTh, st.ilPkMargin, ilPkMax);
  setEng('rsense', rsense, 'Ω');

  setText(
    'eq34_sub',
    `RS = VCS-TH/(m·IL,pk) = ${fmtShort(st.vcsTh)} V/(${fmtShort(st.ilPkMargin)}·${fmtShort(ilPkMax)} A) = ${fmtEng(rsense, 'Ω')}`
  );

  const ilPkSc = eq35_ilPkSc(st.vcsTh, rsense, st.vinMax, st.tDelay, st.lUsed);
  setEng('ilPkSc', ilPkSc, 'A');

  setText(
    'eq35_sub',
    `IL,pk(sc) = VCS-TH/RS + VIN_max·tDelay/L = ${fmtShort(st.vcsTh)}/${fmtShort(rsense)} + ${fmtShort(st.vinMax)}·${fmtEng(st.tDelay, 's')}/${fmtEng(st.lUsed, 'H')} = ${fmtEng(ilPkSc, 'A')}`
  );

  const coutMin = eq36_coutMin(st.lUsed, st.iout, st.vout, st.voutOvershoot);
  setEng('coutMin', coutMin, 'F');

  // Eq.36 substitution line
  const denom36 = (st.vout + st.voutOvershoot) ** 2 - st.vout ** 2;
  setText(
    'eq36_sub',
    `COUT = L·IOUT^2/((VOUT+ΔV)^2−VOUT^2) = ${fmtEng(st.lUsed, 'H')}·${fmtShort(st.iout)}^2/((${fmtShort(st.vout)}+${fmtShort(st.voutOvershoot)})^2−${fmtShort(st.vout)}^2) = ${fmtEng(coutMin, 'F')}`
  );

  const voutRipple = eq37_voutRipple(dIlNom, st.fsw, st.coutEff, st.routEsr);
  setEng('voutRipple', voutRipple, 'Vpp');

  // Eq.37 substitution line
  const vcap37 = dIlNom / (8 * st.fsw * st.coutEff);
  const vesr37 = dIlNom * st.routEsr;
  setText(
    'eq37_sub',
    `ΔVOUT,pp = sqrt((ΔIL/(8·FSW·COUT,eff))^2 + (ΔIL·RESR)^2) = sqrt((${fmtShort(dIlNom)}/(8·${fmtHz(st.fsw)}·${fmtEng(st.coutEff, 'F')}))^2 + (${fmtShort(dIlNom)}·${fmtShort(st.routEsr)})^2) = ${fmtEng(voutRipple, 'V')}`
  );

  const icoRms = eq38_icoRms(dIlNom);
  setEng('icoRms', icoRms, 'A');

  // Eq.38 substitution line
  setText(
    'eq38_sub',
    `ICO,rms = ΔIL/sqrt(12) = ${fmtShort(dIlNom)}/sqrt(12) = ${fmtEng(icoRms, 'A')}`
  );

  const icinRms = eq39_icinRms(st.iout, st.duty);
  setEng('icinRms', icinRms, 'A');

  // Eq.39 substitution line
  setText(
    'eq39_sub',
    `ICIN,rms = IOUT·sqrt(D(1−D)) = ${fmtShort(st.iout)}·sqrt(${fmtShort(st.duty)}·(1−${fmtShort(st.duty)})) = ${fmtEng(icinRms, 'A')}`
  );

  const { cin, note } = eq40_cinRequired(st.iout, st.duty, st.fsw, st.vinRippleSpec, st.rinEsr);
  setText('cinReq', (cin === Infinity) ? '∞' : fmtEng(cin, 'F'));
  setText('cinNote', note || '');

  // Substitution line for Eq.40
  const icinRms40 = eq39_icinRms(st.iout, st.duty);
  const dvEsr40 = icinRms40 * st.rinEsr;
  const dvCapAllow = (dvEsr40 >= st.vinRippleSpec)
    ? NaN
    : Math.sqrt(Math.max(st.vinRippleSpec ** 2 - dvEsr40 ** 2, 0));
  const iFactor = st.iout * st.duty * (1 - st.duty);
  const cinCalc = (Number.isFinite(dvCapAllow) && dvCapAllow > 0)
    ? (iFactor / (st.fsw * dvCapAllow))
    : Infinity;

  setText(
    'eq40_sub',
    `ICIN,rms = IOUT·sqrt(D(1−D)) = ${fmtShort(st.iout)}·sqrt(${fmtShort(st.duty)}·(1−${fmtShort(st.duty)})) = ${fmtEng(icinRms40, 'A')};  ` +
    `ΔVESR = ICIN,rms·RESR = ${fmtShort(icinRms40)}·${fmtShort(st.rinEsr)} = ${fmtEng(dvEsr40, 'V')};  ` +
    `Cin = IOUT·D(1−D)/(FSW·sqrt(ΔVIN^2−ΔVESR^2)) = ${fmtShort(st.iout)}·${fmtShort(st.duty)}·(1−${fmtShort(st.duty)})/(${fmtHz(st.fsw)}·${fmtShort(dvCapAllow)}) = ${(cinCalc === Infinity) ? '∞' : fmtEng(cinCalc, 'F')}`
  );

  const rt = eq41_rtOhm(st.fsw);
  setEng('rt', rt, 'Ω');

  // Eq.41 substitution line (solve for RT)
  setText(
    'eq41_sub',
    `RT = (((10^6)/(FSW[kHz]))−53)/45  [kΩ] = (((10^6)/${fmtShort(st.fsw / 1000)})−53)/45 = ${fmtEng(rt, 'Ω')}`
  );

  const rtop = eq42_rtop(st.vout, st.vref, st.rfbBot);
  setEng('rfb2Val', st.rfbBot, 'Ω');
  setEng('rfb1', rtop, 'Ω');

  // Eq.42 substitution line
  setText(
    'eq42_sub',
    `RFB1 = RFB2·(VOUT/VREF−1) = ${fmtEng(st.rfbBot, 'Ω')}·(${fmtShort(st.vout)}/${fmtShort(st.vref)}−1) = ${fmtEng(rtop, 'Ω')}`
  );

  const rcomp = eq43_rcomp(st.vout, rsense, st.gm, st.fc, st.coutEff, st.vref);
  setEng('rcomp', rcomp, 'Ω');

  // Eq.43 substitution line
  setText(
    'eq43_sub',
    `RCOMP = (VOUT·RS·Gm)/(2π·fC·COUT,eff·VREF) = (${fmtShort(st.vout)}·${fmtShort(rsense)}·${fmtEng(st.gm, 'S')})/(2π·${fmtShort(st.fc)}·${fmtEng(st.coutEff, 'F')}·${fmtShort(st.vref)}) = ${fmtEng(rcomp, 'Ω')}`
  );

  const ccomp = eq44_ccomp(st.fc, rcomp);
  setEng('ccomp', ccomp, 'F');

  // Eq.44 substitution line
  setText(
    'eq44_sub',
    `CCOMP = 10/(2π·fC·RCOMP) = 10/(2π·${fmtShort(st.fc)}·${fmtEng(rcomp, 'Ω')}) = ${fmtEng(ccomp, 'F')}`
  );

  const chf = eq45_chf(st.fesr, rcomp, st.cbw);
  setEng('chf', chf, 'F');

  // Eq.45 substitution line
  setText(
    'eq45_sub',
    `CHF = 1/(2π·fESR·RCOMP) − CBW = 1/(2π·${fmtShort(st.fesr)}·${fmtEng(rcomp, 'Ω')}) − ${fmtEng(st.cbw, 'F')} = ${fmtEng(chf, 'F')}`
  );

  lastResults = {
    deltaIlNom_A: dIlNom,
    lReq_H: Lreq,
    deltaIlMax_A: dIlMax,
    ilPkMax_A: ilPkMax,
    lo_sc_H: loSc,
    // Backward-compat alias (older exports used the key "rosc")
    rosc: loSc,
    rsense_Ohm: rsense,
    ilPkSc_A: ilPkSc,
    coutMin_F: coutMin,
    voutRipple_Vpp: voutRipple,
    icoRms_A: icoRms,
    icinRms_A: icinRms,
    cinReq_F: cin,
    rt_Ohm: rt,
    rfbTop_Ohm: rtop,
    rcomp_Ohm: rcomp,
    ccomp_F: ccomp,
    chf_F: chf,
  };

  updateSchematic(st, lastResults, rsense);

  saveState(st);
}

function buildPayload() {
  const st = readInputs();
  return {
    meta: {
      tool: 'lm5148_webapp',
      version: 1,
      exportedAt: new Date().toISOString(),
    },
    inputs: st,
    results: lastResults,
  };
}

function downloadJson(filename, obj) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function hookEvents() {
  for (const key of Object.keys(DEFAULTS)) {
    const el = $(key);
    if (!el) continue;
    el.addEventListener('input', () => {
      try {
        recalc();
        setStatus('Updated.');
      } catch {
        setStatus('Some inputs are invalid.');
      }
    });
  }

  $('btnReset').addEventListener('click', () => {
    writeInputs(DEFAULTS);
    recalc();
    setStatus('Reset to defaults.');
    if (window.MathJax?.typesetPromise) window.MathJax.typesetPromise();
  });

  $('btnCopy').addEventListener('click', async () => {
    const payload = buildPayload();

    try {
      await navigator.clipboard.writeText(JSON.stringify(payload, null, 2));
      setStatus('Copied results JSON to clipboard.');
    } catch {
      setStatus('Clipboard copy failed (browser permissions).');
    }
  });

  $('btnDownload').addEventListener('click', () => {
    const payload = buildPayload();
    downloadJson('lm5148_design.json', payload);
    setStatus('Downloaded lm5148_design.json');
  });

  const btnApplyUvlo = $('btnApplyUvlo');
  if (btnApplyUvlo) {
    btnApplyUvlo.addEventListener('click', () => {
      try {
        const st = readInputs();
        const uvloR1Calc = uvlo_ruv1(st.uvloVinOn, st.uvloVinOff, st.uvloIhys);
        const uvloR2Calc = uvlo_ruv2(uvloR1Calc, st.uvloVen, st.uvloVinOn);
        if (!Number.isFinite(uvloR1Calc) || !Number.isFinite(uvloR2Calc)) {
          setStatus('UVLO inputs invalid; cannot copy.');
          return;
        }
        if ($('uvloR1')) $('uvloR1').value = String(uvloR1Calc);
        if ($('uvloR2')) $('uvloR2').value = String(uvloR2Calc);
        recalc();
        setStatus('Copied UVLO calc to RUV1/RUV2.');
      } catch {
        setStatus('UVLO copy failed (invalid inputs).');
      }
    });
  }
}

function init() {
  const state = loadState();
  writeInputs(state);
  hookEvents();
  hookSchematicHotspots();
  preferOptionalImages();
  try {
    recalc();
    setStatus('Ready.');
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    setStatus(`Error: ${msg}`);
    // Also log for debugging.
    // eslint-disable-next-line no-console
    console.error(err);
  }

  // Ensure MathJax renders after initial DOM.
  if (window.MathJax?.typesetPromise) {
    window.MathJax.typesetPromise();
  }
}

document.addEventListener('DOMContentLoaded', init);

// Surface unexpected runtime errors in the on-page status line.
window.addEventListener('error', (event) => {
  try {
    const msg = event?.error?.message || event?.message || 'Unknown error';
    setStatus(`Error: ${msg}`);
  } catch {
    // ignore
  }
});
