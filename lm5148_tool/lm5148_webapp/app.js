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

  // Eq33 helper
  rsForEq33: 5e-3,

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
  rfbBot: 10_000,

  // Compensation
  fc: 60_000,
  gm: 1200e-6,
  fesr: 500_000,
  cbw: 0.8e-12,
};

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

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULTS };
    const obj = JSON.parse(raw);
    return { ...DEFAULTS, ...obj };
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
    state[key] = safeNumber(el.value);
  }
  return state;
}

function writeInputs(state) {
  for (const key of Object.keys(DEFAULTS)) {
    const el = $(key);
    if (!el) continue;
    el.value = state[key];
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

function eq33_rosc(vout, rs, L, fsw) {
  return (vout * rs) / (L * fsw);
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

  // Derived values
  const dIlNom = rippleDeltaIl(st.iout, st.rippleFrac);
  $('deltaIlNom').value = Number.isFinite(dIlNom) ? dIlNom : '';

  const Lreq = eq31_L(st.vinNom, st.vout, st.fsw, dIlNom);
  setEng('lReq', Lreq, 'H');

  const dIlMax = deltaIl(st.vinMax, st.vout, st.fsw, st.lUsed);
  setEng('deltaIlMax', dIlMax, 'A');

  const ilPkMax = eq32_ilPk(st.iout, dIlMax);
  setEng('ilPkMax', ilPkMax, 'A');

  const rosc = eq33_rosc(st.vout, st.rsForEq33, st.lUsed, st.fsw);
  setText('rosc', fmt(rosc));

  const rsense = eq34_rsense(st.vcsTh, st.ilPkMargin, ilPkMax);
  setEng('rsense', rsense, 'Ω');

  const ilPkSc = eq35_ilPkSc(st.vcsTh, rsense, st.vinMax, st.tDelay, st.lUsed);
  setEng('ilPkSc', ilPkSc, 'A');

  const coutMin = eq36_coutMin(st.lUsed, st.iout, st.vout, st.voutOvershoot);
  setEng('coutMin', coutMin, 'F');

  const voutRipple = eq37_voutRipple(dIlNom, st.fsw, st.coutEff, st.routEsr);
  setEng('voutRipple', voutRipple, 'Vpp');

  const icoRms = eq38_icoRms(dIlNom);
  setEng('icoRms', icoRms, 'A');

  const icinRms = eq39_icinRms(st.iout, st.duty);
  setEng('icinRms', icinRms, 'A');

  const { cin, note } = eq40_cinRequired(st.iout, st.duty, st.fsw, st.vinRippleSpec, st.rinEsr);
  setText('cinReq', (cin === Infinity) ? '∞' : fmtEng(cin, 'F'));
  setText('cinNote', note || '');

  const rt = eq41_rtOhm(st.fsw);
  setEng('rt', rt, 'Ω');

  const rtop = eq42_rtop(st.vout, st.vref, st.rfbBot);
  setEng('rfbTop', rtop, 'Ω');

  const rcomp = eq43_rcomp(st.vout, rsense, st.gm, st.fc, st.coutEff, st.vref);
  setEng('rcomp', rcomp, 'Ω');

  const ccomp = eq44_ccomp(st.fc, rcomp);
  setEng('ccomp', ccomp, 'F');

  const chf = eq45_chf(st.fesr, rcomp, st.cbw);
  setEng('chf', chf, 'F');

  lastResults = {
    deltaIlNom_A: dIlNom,
    lReq_H: Lreq,
    deltaIlMax_A: dIlMax,
    ilPkMax_A: ilPkMax,
    rosc: rosc,
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
}

function init() {
  const state = loadState();
  writeInputs(state);
  hookEvents();
  recalc();
  setStatus('Ready.');

  // Ensure MathJax renders after initial DOM.
  if (window.MathJax?.typesetPromise) {
    window.MathJax.typesetPromise();
  }
}

document.addEventListener('DOMContentLoaded', init);
