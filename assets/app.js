/* Dashboard de contratos de crédito público · Gobernación del Valle del Cauca
 * Lee la pestaña «ANDANDO» del Excel (data/contratos.xlsx o el archivo que cargue el usuario),
 * calcula indicadores y renderiza gráficas, tabla y reportes PDF/Excel en el navegador. */
(() => {
  'use strict';

  const XLSX = window.XLSX;
  const ChartJS = window.Chart;
  const { jsPDF } = window.jspdf;

  const DATA_URL = 'data/contratos.xlsx';
  const SHEET_NAME = 'ANDANDO';
  const DAY = 86400000;
  const YEAR_DAYS = 365.25;
  const TODAY = (() => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); })();
  const FONT = 'system-ui, -apple-system, "Segoe UI", Roboto, sans-serif';
  const LS_THEME = 'cgv-theme';
  const LS_COLS = 'cgv-columns-v1';

  /* ------------------------------------------------------------------ formato */
  const nf0 = new Intl.NumberFormat('es-CO', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('es-CO', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nf2 = new Intl.NumberFormat('es-CO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const pct = (v) => (Number.isFinite(v) ? nf1.format(v * 100) + ' %' : '—');
  const fmtCOP = (v) => (Number.isFinite(v) ? '$ ' + nf0.format(v) : '—');
  const fmtMilM = (v) => (Number.isFinite(v) ? '$ ' + nf1.format(v / 1e9) + ' mil M' : '—');
  const fmtBig = (v) => {
    if (!Number.isFinite(v)) return '—';
    if (Math.abs(v) >= 1e12) return '$ ' + nf2.format(v / 1e12) + ' billones';
    if (Math.abs(v) >= 1e9) return '$ ' + nf1.format(v / 1e9) + ' mil M';
    return '$ ' + nf0.format(v / 1e6) + ' M';
  };
  const fmtYears = (v) => (Number.isFinite(v) ? nf1.format(v) + ' años' : '—');
  const fmtDate = (d) => (d instanceof Date ? String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear() : 'Sin fecha');
  const norm = (s) => String(s ?? '').normalize('NFKD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
  const sentence = (s) => { const t = String(s ?? '').trim(); return t ? t.charAt(0).toUpperCase() + t.slice(1).toLowerCase() : 'Sin estado'; };

  /* ------------------------------------------------------------- DOM helpers */
  const $ = (sel, root = document) => root.querySelector(sel);
  function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    if (attrs) {
      for (const [k, v] of Object.entries(attrs)) {
        if (v == null || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'style') el.style.cssText = v;
        else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
        else el.setAttribute(k, v === true ? '' : v);
      }
    }
    for (const kid of kids.flat()) {
      if (kid == null || kid === false) continue;
      el.append(kid instanceof Node ? kid : document.createTextNode(String(kid)));
    }
    return el;
  }
  const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
  const store = {
    get(k) { try { return localStorage.getItem(k); } catch { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch { /* sin almacenamiento */ } },
  };

  /* --------------------------------------------------------- normalización */
  const TYPE_PUBLIC = 'Banca pública y de fomento';
  const TYPE_PRIVATE = 'Banca privada';
  const TYPE_OTHER = 'Sin clasificar';
  const LENDERS = [
    [/financiera de desarrollo territorial|findeter/, 'Findeter', TYPE_PUBLIC],
    [/infivalle|instituto financiero para el desarrollo del valle/, 'Infivalle', TYPE_PUBLIC],
    [/agrario/, 'Banco Agrario', TYPE_PUBLIC],
    [/bilbao|bbva/, 'BBVA', TYPE_PRIVATE],
    [/davivienda/, 'Davivienda', TYPE_PRIVATE],
    [/occidente/, 'Banco de Occidente', TYPE_PRIVATE],
    [/popular/, 'Banco Popular', TYPE_PRIVATE],
    [/av villas/, 'Banco AV Villas', TYPE_PRIVATE],
    [/bbogota|banco de bogota/, 'Banco de Bogotá', TYPE_PRIVATE],
    [/no habilitada/, 'Cuenta no habilitada', TYPE_OTHER],
  ];
  function classifyLender(raw) {
    const n = norm(raw);
    for (const [re, short, type] of LENDERS) if (re.test(n)) return { short, type };
    const t = String(raw || 'Sin acreedor').toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase());
    return { short: t, type: TYPE_OTHER };
  }
  function periodOf(d) {
    if (!d) return 'Sin fecha';
    const start = 2020 + 4 * Math.floor((d.getFullYear() - 2020) / 4);
    return `Administración ${start}–${start + 3}`;
  }

  const COLS = {
    entity: /^nombre entidad/, ctype: /^tipo de contrato/, lender: /^proveedor adjudicado/, nit: /^documento proveedor/,
    object: /^objeto/, value: /^valor del contrato/, status: /^estado/, ref: /^referencia/, signed: /^fecha de firma/,
    start: /^fecha de inicio/, end: /^fecha de fin/, mods: /^n\W*o?\W*de modificaciones/, modTypes: /^tipos de modificacion/,
    valueAdds: /adiciones en valor/, daysAdded: /dias adicionados/, url: /^url/,
  };

  function toDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v)) { const t = new Date(v.getTime() + 3600e3); return new Date(t.getFullYear(), t.getMonth(), t.getDate()); }
    if (typeof v === 'number') { const p = XLSX.SSF.parse_date_code(v); return p ? new Date(p.y, p.m - 1, p.d) : null; }
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    return null;
  }
  function toNum(v) {
    if (typeof v === 'number') return v;
    if (v == null) return NaN;
    const s = String(v).replace(/[^\d,.-]/g, '');
    if (!s) return NaN;
    // "35.340.551.465,00" → miles con punto (formato colombiano)
    const cleaned = /,\d{1,2}$/.test(s) ? s.replace(/\./g, '').replace(',', '.') : s.replace(/[.,](?=\d{3}(\D|$))/g, '');
    return Number(cleaned);
  }
  const shortRef = (ref) => { const parts = String(ref || '').split(/[-.]/); return parts[parts.length - 1] || String(ref || '—'); };

  function parseSheet(ws) {
    const raw = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
    const out = [];
    raw.forEach((obj, i) => {
      const r = {};
      for (const [k, v] of Object.entries(obj)) {
        const nk = norm(k);
        for (const [field, re] of Object.entries(COLS)) if (r[field] === undefined && re.test(nk)) { r[field] = v; break; }
      }
      if (r.lender == null && r.value == null && r.ref == null) return; // fila vacía
      const lenderRaw = String(r.lender ?? '').trim();
      const { short, type } = classifyLender(lenderRaw);
      const signed = toDate(r.signed), start = toDate(r.start), end = toDate(r.end);
      const value = toNum(r.value);
      const term = start && end && end > start ? (end - start) / DAY / YEAR_DAYS : NaN;
      const rec = {
        id: i, row: i + 2, entity: r.entity ?? '', ctype: r.ctype ?? '',
        lenderRaw, lender: short, ltype: type, nit: r.nit != null ? String(r.nit) : '',
        object: String(r.object ?? '').replace(/\s+/g, ' ').trim(),
        value, status: sentence(r.status), ref: String(r.ref ?? '').trim() || '—',
        signed, start, end, term,
        remaining: end ? Math.max(0, (end - TODAY) / DAY / YEAR_DAYS) : NaN,
        annual: Number.isFinite(term) && term > 0 ? value / term : NaN,
        expired: !!(end && end < TODAY),
        mods: Number.isFinite(toNum(r.mods)) ? toNum(r.mods) : 0,
        modTypes: r.modTypes ? String(r.modTypes) : '',
        valueAdds: Number.isFinite(toNum(r.valueAdds)) ? toNum(r.valueAdds) : 0,
        daysAdded: Number.isFinite(toNum(r.daysAdded)) ? toNum(r.daysAdded) : 0,
        url: r.url ? String(r.url).trim() : '',
        signYear: (signed || start) ? (signed || start).getFullYear() : null,
        period: periodOf(signed || start),
        flags: [],
      };
      rec.shortRef = shortRef(rec.ref);
      rec.label = `${rec.lender} · ${rec.shortRef}${rec.signYear ? ` (${rec.signYear})` : ''}`;
      rec.labelShort = `${rec.lender.replace(/^Banco (de )?/, '')} · ${rec.shortRef.slice(-5)}${rec.signYear ? ` '${String(rec.signYear).slice(2)}` : ''}`;
      if (!start) rec.flags.push('sin fecha de inicio');
      if (/no habilitada/i.test(lenderRaw)) rec.flags.push('proveedor «cuenta no habilitada»');
      if (!Number.isFinite(value) || value <= 0) rec.flags.push('sin valor');
      out.push(rec);
    });
    // Duplicados: un registro con inconsistencias que tiene el mismo valor y fecha de fin que un contrato válido
    // (p. ej. «CUENTA NO HABILITADA» 1.120.19.7-20822 = BBVA 1.120.19.07-20822) se elimina del análisis.
    const removed = [];
    for (const r of out) {
      if (!r.flags.length) continue;
      const twin = out.find((o) => o !== r && !o.flags.length && o.value === r.value && o.end && r.end && +o.end === +r.end);
      if (twin) removed.push(r);
    }
    return { records: out.filter((r) => !removed.includes(r)), removed };
  }

  /* ----------------------------------------------------------------- estado */
  const BUCKETS = [
    { lo: 0, hi: 5, label: 'Menos de 5 años' },
    { lo: 5, hi: 7, label: '5 a 7 años' },
    { lo: 7, hi: 9, label: '7 a 9 años' },
    { lo: 9, hi: 10.5, label: '9 a 10,5 años (≈10)' },
    { lo: 10.5, hi: Infinity, label: 'Más de 10,5 años' },
  ];
  const bucketOf = (t) => (Number.isFinite(t) ? BUCKETS.findIndex((b) => t >= b.lo && t < b.hi) : -1);

  const DIM_LABEL = { period: 'Administración', ltype: 'Tipo de acreedor', status: 'Estado', lender: 'Acreedor' };
  const blankFilters = () => ({
    yearFrom: null, yearTo: null, lender: new Set(), ltype: new Set(), period: new Set(), status: new Set(),
    termMin: null, termMax: null, amtMin: null, amtMax: null, q: '', bucket: null, includeFlagged: false,
  });
  const state = {
    all: [], fileName: DATA_URL, sheetName: SHEET_NAME, filters: blankFilters(), colorBy: 'period',
    scatterScale: 'linear', hidden: { scatter: new Set(), years: new Set(), gantt: new Set(), amort: new Set() },
    colorMaps: {}, rows: [], specs: {}, kpis: [], corr: null, insight: [],
    sort: { key: 'start', dir: 1 }, colFilters: {}, columns: null,
  };

  /* ------------------------------------------------------------- tema/colores */
  let T = {};
  function readTokens() {
    const cs = getComputedStyle(document.documentElement);
    const v = (n) => cs.getPropertyValue(n).trim();
    T = {
      text: v('--text'), text2: v('--text-2'), muted: v('--muted'), grid: v('--grid'), axis: v('--axis'),
      surface: v('--surface'), surface2: v('--surface-2'), border: v('--border'), s: [v('--s1'), v('--s2'), v('--s3')],
      neutral: v('--neutral-mark'), trend: v('--trend'), today: v('--today'), accent: v('--accent'),
    };
  }
  function alpha(hex, a) {
    const m = hex.replace('#', '');
    const n = parseInt(m.length === 3 ? m.split('').map((c) => c + c).join('') : m, 16);
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
  }
  const isDark = () => {
    const a = document.documentElement.getAttribute('data-theme');
    return a ? a === 'dark' : window.matchMedia('(prefers-color-scheme: dark)').matches;
  };

  // Colores asignados a la entidad (no al rango): se calculan sobre todo el archivo, máximo 3 categorías + «Otros».
  function buildColorMaps() {
    const orders = {
      period: [...new Set(state.all.map((r) => r.period))].sort((a, b) => (a === 'Sin fecha') - (b === 'Sin fecha') || a.localeCompare(b)),
      ltype: [TYPE_PUBLIC, TYPE_PRIVATE, TYPE_OTHER].filter((t) => state.all.some((r) => r.ltype === t)),
      status: countBy(state.all, (r) => r.status).map(([k]) => k),
    };
    state.colorMaps = {};
    for (const [dim, cats] of Object.entries(orders)) {
      const map = new Map();
      cats.forEach((c, i) => map.set(c, i < 3 ? i : -1));
      state.colorMaps[dim] = map;
    }
  }
  const catOf = (r, dim = state.colorBy) => { const i = state.colorMaps[dim]?.get(r[dim]); return i === -1 ? 'Otros' : r[dim]; };
  const colorOfCat = (cat, dim = state.colorBy) => { const i = state.colorMaps[dim]?.get(cat); return i == null || i < 0 ? T.neutral : T.s[i]; };
  const catsOf = (rows, dim = state.colorBy) => {
    const present = new Set(rows.map((r) => catOf(r, dim)));
    const ordered = [...state.colorMaps[dim].keys()].filter((c) => present.has(c));
    if (present.has('Otros')) ordered.push('Otros');
    return ordered;
  };

  /* --------------------------------------------------------------- utilidades */
  function countBy(rows, fn) {
    const m = new Map();
    rows.forEach((r) => m.set(fn(r), (m.get(fn(r)) || 0) + 1));
    return [...m.entries()].sort((a, b) => b[1] - a[1]);
  }
  function groupSum(rows, keyFn, valFn = (r) => r.value) {
    const m = new Map();
    rows.forEach((r) => { const k = keyFn(r); const g = m.get(k) || { key: k, sum: 0, n: 0, rows: [] }; g.sum += valFn(r) || 0; g.n++; g.rows.push(r); m.set(k, g); });
    return [...m.values()];
  }
  const sum = (arr, fn = (x) => x) => arr.reduce((a, x) => a + (Number(fn(x)) || 0), 0);
  const wavg = (rows, fn, wfn = (r) => r.value) => {
    const ok = rows.filter((r) => Number.isFinite(fn(r)) && Number.isFinite(wfn(r)));
    const w = sum(ok, wfn);
    return w ? sum(ok, (r) => fn(r) * wfn(r)) / w : NaN;
  };
  const median = (xs) => { const s = xs.filter(Number.isFinite).sort((a, b) => a - b); if (!s.length) return NaN; const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
  function pearson(xs, ys) {
    const n = xs.length; if (n < 3) return NaN;
    const mx = sum(xs) / n, my = sum(ys) / n;
    let sxy = 0, sxx = 0, syy = 0;
    for (let i = 0; i < n; i++) { sxy += (xs[i] - mx) * (ys[i] - my); sxx += (xs[i] - mx) ** 2; syy += (ys[i] - my) ** 2; }
    return sxx && syy ? sxy / Math.sqrt(sxx * syy) : NaN;
  }
  function ranks(xs) {
    const idx = xs.map((v, i) => [v, i]).sort((a, b) => a[0] - b[0]);
    const r = new Array(xs.length);
    for (let i = 0; i < idx.length;) {
      let j = i; while (j + 1 < idx.length && idx[j + 1][0] === idx[i][0]) j++;
      const avg = (i + j) / 2 + 1; for (let k = i; k <= j; k++) r[idx[k][1]] = avg; i = j + 1;
    }
    return r;
  }
  function ols(xs, ys) {
    const n = xs.length; if (n < 2) return null;
    const mx = sum(xs) / n, my = sum(ys) / n;
    let sxy = 0, sxx = 0;
    for (let i = 0; i < n; i++) { sxy += (xs[i] - mx) * (ys[i] - my); sxx += (xs[i] - mx) ** 2; }
    if (!sxx) return null;
    const b = sxy / sxx; return { a: my - b * mx, b };
  }
  const strength = (r) => { const a = Math.abs(r); return a < 0.2 ? 'muy débil' : a < 0.4 ? 'débil' : a < 0.6 ? 'moderada' : a < 0.8 ? 'fuerte' : 'muy fuerte'; };
  const jan1 = (y) => new Date(y, 0, 1);
  const yearFrac = (d) => (d - jan1(d.getFullYear())) / (jan1(d.getFullYear() + 1) - jan1(d.getFullYear()));
  const range = (a, b) => { const o = []; for (let i = a; i <= b; i++) o.push(i); return o; };

  /* ----------------------------------------------------------------- filtrado */
  function filtered(except) {
    const f = state.filters;
    const q = norm(f.q);
    return state.all.filter((r) => {
      if (!f.includeFlagged && r.flags.length) return false;
      if (except !== 'year') {
        if (f.yearFrom != null && !(r.signYear >= f.yearFrom)) return false;
        if (f.yearTo != null && !(r.signYear <= f.yearTo)) return false;
      }
      if (except !== 'lender' && f.lender.size && !f.lender.has(r.lender)) return false;
      if (except !== 'ltype' && f.ltype.size && !f.ltype.has(r.ltype)) return false;
      if (f.period.size && !f.period.has(r.period)) return false;
      if (f.status.size && !f.status.has(r.status)) return false;
      if (f.termMin != null && !(r.term >= f.termMin)) return false;
      if (f.termMax != null && !(r.term <= f.termMax)) return false;
      if (f.amtMin != null && !(r.value >= f.amtMin * 1e9)) return false;
      if (f.amtMax != null && !(r.value <= f.amtMax * 1e9)) return false;
      if (except !== 'bucket' && f.bucket != null && bucketOf(r.term) !== f.bucket) return false;
      if (q && !norm(`${r.ref} ${r.lender} ${r.lenderRaw} ${r.object} ${r.nit} ${r.status}`).includes(q)) return false;
      return true;
    });
  }

  /* ---------------------------------------------------------- Chart.js base */
  ChartJS.defaults.font.family = FONT;
  ChartJS.defaults.font.size = 12;
  ChartJS.defaults.animation.duration = 300;
  ChartJS.defaults.maintainAspectRatio = false;
  ChartJS.defaults.devicePixelRatio = Math.max(window.devicePixelRatio || 1, 2);

  ChartJS.register({
    id: 'surfaceBg',
    beforeDraw(c) { const { ctx } = c; ctx.save(); ctx.globalCompositeOperation = 'destination-over'; ctx.fillStyle = T.surface; ctx.fillRect(0, 0, c.width, c.height); ctx.restore(); },
  });
  // Etiqueta al final de barras (total de la pila en barras apiladas)
  ChartJS.register({
    id: 'barLabels',
    afterDatasetsDraw(c, _a, o) {
      if (!o || !o.formatter) return;
      const { ctx } = c; const horiz = c.options.indexAxis === 'y';
      const n = c.data.labels.length;
      ctx.save(); ctx.font = `500 11px ${FONT}`; ctx.fillStyle = T.text2;
      for (let i = 0; i < n; i++) {
        const txt = o.formatter(i); if (txt == null) continue;
        let el = null;
        c.data.datasets.forEach((ds, di) => {
          if (ds.type && ds.type !== 'bar') return;
          const meta = c.getDatasetMeta(di); if (meta.hidden || !c.isDatasetVisible(di)) return;
          const e = meta.data[i]; const v = ds.data[i];
          if (!e || !v) return;
          if (!el || (horiz ? e.x > el.x : e.y < el.y)) el = e;
        });
        if (!el) continue;
        if (horiz) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(txt, el.x + 6, el.y); }
        else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(txt, el.x, el.y - 4); }
      }
      ctx.restore();
    },
  });
  // Línea de referencia vertical («Hoy»)
  ChartJS.register({
    id: 'refLine',
    afterDatasetsDraw(c, _a, o) {
      if (!o || o.value == null || !c.scales.x) return;
      const x = c.scales.x.getPixelForValue(o.value);
      const { top, bottom, left, right } = c.chartArea;
      if (x < left || x > right) return;
      const { ctx } = c; ctx.save();
      ctx.strokeStyle = T.today; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.moveTo(x, top); ctx.lineTo(x, bottom); ctx.stroke();
      ctx.fillStyle = T.today; ctx.font = `600 11px ${FONT}`; ctx.textBaseline = 'top';
      ctx.textAlign = x > right - 40 ? 'right' : 'left'; ctx.fillText(o.label || 'Hoy', x + (x > right - 40 ? -4 : 4), top + 2);
      ctx.restore();
    },
  });
  // Etiquetas directas selectivas para puntos
  ChartJS.register({
    id: 'pointLabels',
    afterDatasetsDraw(c, _a, o) {
      if (!o || !o.items) return;
      const { ctx } = c; ctx.save(); ctx.font = `500 11px ${FONT}`; ctx.fillStyle = T.text2; ctx.textBaseline = 'middle';
      for (const it of o.items) {
        if (!c.isDatasetVisible(it.ds)) continue;
        const el = c.getDatasetMeta(it.ds).data[it.i]; if (!el) continue;
        const right = el.x + 10 + ctx.measureText(it.text).width < c.chartArea.right;
        ctx.textAlign = right ? 'left' : 'right'; ctx.fillText(it.text, el.x + (right ? 10 : -10), el.y);
      }
      ctx.restore();
    },
  });
  // Mira vertical en gráficas de línea
  ChartJS.register({
    id: 'crosshair',
    afterDatasetsDraw(c, _a, o) {
      if (!o || !o.enabled) return;
      const act = c.tooltip && c.tooltip.getActiveElements(); if (!act || !act.length) return;
      const x = act[0].element.x; const { top, bottom } = c.chartArea; const { ctx } = c;
      ctx.save(); ctx.strokeStyle = T.axis; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(x, top); ctx.lineTo(x, bottom); ctx.stroke(); ctx.restore();
    },
  });

  function tooltipBase(extra = {}) {
    return {
      backgroundColor: T.surface, borderColor: T.axis, borderWidth: 1, titleColor: T.text2, bodyColor: T.text, footerColor: T.muted,
      titleFont: { weight: '500', size: 12 }, bodyFont: { weight: '600', size: 12 }, footerFont: { weight: '400', size: 11 },
      padding: 10, cornerRadius: 8, boxWidth: 12, boxHeight: 2, boxPadding: 6, caretSize: 5,
      callbacks: {
        labelColor(ctx) {
          const ds = ctx.dataset; let col = ds.keyColor || ds.backgroundColor;
          if (typeof col === 'function') col = ds.keyColor || T.s[0];
          if (Array.isArray(col)) col = col[ctx.dataIndex];
          return { borderColor: col, backgroundColor: col, borderWidth: 0 };
        },
        ...extra,
      },
    };
  }
  const axis = (o = {}) => ({
    grid: { color: T.grid, lineWidth: 1, drawTicks: false, ...(o.grid || {}) },
    border: { color: T.axis, display: true },
    ticks: { color: T.muted, padding: 6, font: { size: 11 }, ...(o.ticks || {}) },
    title: o.title ? { display: true, text: o.title, color: T.text2, font: { size: 11, weight: '500' } } : { display: false },
    ...Object.fromEntries(Object.entries(o).filter(([k]) => !['grid', 'ticks', 'title'].includes(k))),
  });
  const milM = (v) => nf0.format(v);

  const charts = {};
  function upsert(id, cfg, mode) {
    const canvas = document.getElementById('c-' + id);
    if (!canvas) return;
    const c = charts[id];
    if (c && c.config.type === cfg.type) { c.data = cfg.data; c.options = cfg.options; c.update(mode); return; }
    if (c) c.destroy();
    charts[id] = new ChartJS(canvas, cfg);
  }
  const setBoxHeight = (id, px) => { const b = document.getElementById('box-' + id); if (b) b.style.height = px + 'px'; };
  const compact = () => window.innerWidth < 600;

  function renderLegend(elId, items, hiddenSet) {
    const el = document.getElementById(elId); if (!el) return;
    el.replaceChildren(...items.map((it) => {
      if (it.static) return h('span', { class: 'key' }, h('i', { class: it.shape || '', style: it.color ? `background:${it.color}` : null }), it.label);
      const off = hiddenSet && hiddenSet.has(it.label);
      return h('button', {
        type: 'button', class: off ? 'off' : null, 'aria-pressed': off ? 'false' : 'true', title: 'Mostrar/ocultar',
        onclick: () => { if (!hiddenSet) return; off ? hiddenSet.delete(it.label) : hiddenSet.add(it.label); render('none'); },
      }, h('i', { class: it.shape || '', style: `background:${it.color}` }), it.label);
    }));
  }

  /* ---------------------------------------------------- tablas de cada gráfica */
  // Cada gráfica publica su tabla equivalente (accesibilidad + exportación PDF/Excel).
  function setSpec(id, title, columns, rows, opts = {}) { state.specs[id] = { title, columns, rows, ...opts }; }
  function renderSpecToggles() {
    document.querySelectorAll('[data-chart]').forEach((card) => {
      const id = card.dataset.chart; const spec = state.specs[id]; if (!spec) return;
      let tools = card.querySelector('.card-tools');
      if (!tools) {
        tools = h('div', { class: 'card-tools' });
        const btn = h('button', { type: 'button', class: 'btn btn-ghost btn-sm', 'aria-expanded': 'false' }, 'Ver datos');
        const wrap = h('div', { class: 'mini-wrap', hidden: true });
        btn.addEventListener('click', () => { const open = wrap.hidden; wrap.hidden = !open; btn.setAttribute('aria-expanded', String(open)); btn.textContent = open ? 'Ocultar datos' : 'Ver datos'; });
        tools.append(btn); card.append(tools, wrap);
      }
      const wrap = card.querySelector('.mini-wrap');
      const fmtCell = (v, c) => (c.fmt ? c.fmt(v) : v);
      wrap.replaceChildren(h('table', { class: 'mini-table' },
        h('thead', null, h('tr', null, spec.columns.map((c) => h('th', { class: c.num ? 'num' : null }, c.label)))),
        h('tbody', null, spec.rows.map((r) => h('tr', null, spec.columns.map((c, i) => h('td', { class: c.num ? 'num' : null }, fmtCell(r[i], c))))))));
    });
  }

  /* ------------------------------------------------------------------ KPIs */
  function renderKPIs(rows) {
    const total = sum(rows, (r) => r.value);
    const lenders = new Set(rows.map((r) => r.lender));
    const wTerm = wavg(rows, (r) => r.term), sTerm = sum(rows.filter((r) => Number.isFinite(r.term)), (r) => r.term) / (rows.filter((r) => Number.isFinite(r.term)).length || 1);
    const wRem = wavg(rows, (r) => r.remaining);
    const expired = rows.filter((r) => r.expired);
    const modified = rows.filter((r) => r.mods > 0);
    const biggest = rows.reduce((a, r) => (!a || r.value > a.value ? r : a), null);
    const balToday = balanceAt(rows, TODAY);
    const k = [
      { info: 'total', label: 'Monto total contratado', value: fmtBig(total), note: `${rows.length} contratos · ${lenders.size} acreedores`, hero: true, raw: total },
      { info: 'wterm', label: 'Plazo promedio ponderado', value: fmtYears(wTerm), note: `Promedio simple: ${fmtYears(sTerm)}`, raw: wTerm },
      { info: 'wrem', label: 'Vida remanente ponderada', value: fmtYears(wRem), note: expired.length ? `${expired.length} con fecha fin ya cumplida` : 'Ninguno vencido por fecha fin', raw: wRem },
      { info: 'balance', label: 'Saldo teórico hoy', value: fmtBig(balToday), note: `${pct(total ? balToday / total : NaN)} del contratado (supuesto lineal)`, raw: balToday },
      { info: 'avg', label: 'Monto promedio por contrato', value: fmtBig(rows.length ? total / rows.length : NaN), note: `Mediana: ${fmtBig(median(rows.map((r) => r.value)))}`, raw: rows.length ? total / rows.length : NaN },
      { info: 'mods', label: 'Contratos modificados', value: `${modified.length} de ${rows.length}`, note: `${pct(rows.length ? modified.length / rows.length : NaN)} · ${nf0.format(sum(rows, (r) => r.mods))} modificaciones`, raw: modified.length },
      { info: 'biggest', label: 'Mayor contrato', value: biggest ? fmtBig(biggest.value) : '—', note: biggest ? `${biggest.lender} · ${biggest.ref}` : '', raw: biggest ? biggest.value : NaN },
    ];
    state.kpis = k;
    // Desglose del total por la dimensión de color activa
    const split = catsOf(rows).map((c) => { const rs = rows.filter((r) => catOf(r) === c); return { c, n: rs.length, v: sum(rs, (r) => r.value) }; });
    const splitEl = total ? h('div', { class: 'hero-split' },
      h('div', { class: 'hero-bar' }, split.map((x) => h('span', { style: `width:${(x.v / total) * 100}%;background:${colorOfCat(x.c)}`, title: `${x.c}: ${fmtCOP(x.v)}` }))),
      h('ul', { class: 'hero-legend' }, split.map((x) => h('li', null, h('i', { style: `background:${colorOfCat(x.c)}` }), h('span', null, x.c), h('strong', null, `${fmtBig(x.v)}`), h('span', { class: 'muted' }, ` · ${pct(x.v / total)} · ${x.n}`))))) : null;
    $('#kpis').replaceChildren(...k.map((x) => h('div', { class: 'kpi' + (x.hero ? ' kpi-hero' : '') },
      h('div', null, h('div', { class: 'kpi-label' }, x.label, infoBtn(x.info)), h('div', { class: 'kpi-value' }, x.value)),
      h('div', { class: 'kpi-note' }, x.note), x.hero ? splitEl : null)));
  }

  /* ----------------------------------------------- plazo vs. monto (clave) */
  function renderCorrelation(rows) {
    const pts = rows.filter((r) => Number.isFinite(r.term) && Number.isFinite(r.value));
    const xs = pts.map((r) => r.term), ys = pts.map((r) => r.value / 1e9);
    const r = pearson(xs, ys), rho = pearson(ranks(xs), ranks(ys)), fit = ols(xs, ys);
    state.corr = { n: pts.length, r, r2: r * r, rho, fit };
    const stats = [
      { info: 'pearson', label: 'Correlación de Pearson (r)', value: Number.isFinite(r) ? nf2.format(r) : '—', note: Number.isFinite(r) ? `Relación lineal ${strength(r)}` : 'Se requieren ≥ 3 contratos' },
      { info: 'r2', label: 'R² del ajuste lineal', value: Number.isFinite(r) ? pct(r * r) : '—', note: 'Varianza del monto explicada por el plazo' },
      { info: 'spearman', label: 'Correlación de Spearman (ρ)', value: Number.isFinite(rho) ? nf2.format(rho) : '—', note: 'Por rangos; robusta a valores extremos' },
      { info: 'slope', label: 'Pendiente', value: fit ? `${fit.b >= 0 ? '+' : ''}${nf1.format(fit.b)} mil M` : '—', note: 'Monto adicional por cada año más de plazo' },
      { info: 'annualTotal', label: 'Carga anual implícita total', value: fmtBig(sum(pts, (p) => p.annual)), note: 'Σ monto ÷ plazo de los contratos' },
    ];
    $('#corr-stats').replaceChildren(...stats.map((s) => h('div', { class: 'stat' }, h('div', { class: 'kpi-label' }, s.label, infoBtn(s.info)), h('div', { class: 'kpi-value' }, s.value), h('div', { class: 'kpi-note' }, s.note))));

    // Narrativa automática
    const ins = [];
    if (pts.length >= 3 && Number.isFinite(r)) {
      const dir = r > 0 ? 'positiva' : 'negativa';
      ins.push(`Con los filtros actuales (${pts.length} contratos), la relación entre plazo y monto es ${strength(r)} y ${dir} (r = ${nf2.format(r)}; ρ = ${Number.isFinite(rho) ? nf2.format(rho) : '—'}). ` +
        (Math.abs(r) >= 0.4 ? (r > 0 ? 'Los contratos de mayor monto tienden a pactarse a plazos más largos.' : 'Los contratos de mayor monto tienden a pactarse a plazos más cortos.') : 'El plazo explica poco de la variación del monto: la mayoría de créditos se pacta a ≈10 años con montos muy distintos.') +
        (fit ? ` En promedio, cada año adicional de plazo se asocia con ${fit.b >= 0 ? '+' : '−'}${fmtBig(Math.abs(fit.b) * 1e9)}.` : ''));
      const long = pts.filter((p) => p.term >= 9.5), short = pts.filter((p) => p.term < 9.5);
      const tot = sum(pts, (p) => p.value);
      if (long.length && short.length) {
        ins.push(`Los contratos a ≈10 años o más (${long.length}) concentran el ${pct(sum(long, (p) => p.value) / tot)} del monto, con un promedio de ${fmtBig(sum(long, (p) => p.value) / long.length)} por contrato; los de menos de 9,5 años (${short.length}) promedian ${fmtBig(sum(short, (p) => p.value) / short.length)}.`);
      }
      const top = pts.reduce((a, p) => (!a || p.annual > a.annual ? p : a), null);
      if (top) ins.push(`La mayor carga anual implícita es la de ${top.lender} (${top.ref}): ${fmtBig(top.value)} a ${fmtYears(top.term)} ⇒ ${fmtBig(top.annual)} por año.`);
    } else {
      ins.push('No hay suficientes contratos con fechas válidas para estimar la relación plazo–monto con los filtros actuales.');
    }
    state.insight = ins;
    $('#corr-insight').replaceChildren(...ins.map((t) => h('p', null, t)));

    // Dispersión
    const cats = catsOf(pts);
    const datasets = cats.map((cat) => {
      const col = colorOfCat(cat);
      const data = pts.filter((p) => catOf(p) === cat).map((p) => ({ x: p.term, y: p.value / 1e9, rec: p }));
      return { type: 'scatter', label: cat, data, backgroundColor: col, keyColor: col, borderColor: T.surface, borderWidth: 2, pointRadius: compact() ? 5 : 6, pointHoverRadius: 8, hitRadius: 10, hidden: state.hidden.scatter.has(cat) };
    });
    if (fit && pts.length >= 3) {
      const x0 = Math.min(...xs), x1 = Math.max(...xs);
      datasets.push({ type: 'line', label: 'Tendencia lineal', data: [{ x: x0, y: fit.a + fit.b * x0 }, { x: x1, y: fit.a + fit.b * x1 }], borderColor: T.trend, keyColor: T.trend, borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, pointHitRadius: 0, fill: false, order: 10 });
    }
    // Etiquetas directas: mayor monto y mayor plazo
    const items = [];
    const pick = (fn) => pts.reduce((a, p) => (!a || fn(p) > fn(a) ? p : a), null);
    [pick((p) => p.value), pick((p) => p.term)].filter(Boolean).forEach((p, k, arr) => {
      if (k === 1 && p === arr[0]) return;
      const ds = cats.indexOf(catOf(p)); const i = datasets[ds].data.findIndex((d) => d.rec === p);
      items.push({ ds, i, text: p.lender });
    });
    const logScale = state.scatterScale === 'logarithmic';
    upsert('scatter', {
      type: 'scatter',
      data: { datasets },
      options: {
        layout: { padding: { top: 8, right: 8 } },
        scales: {
          x: axis({ type: 'linear', title: 'Plazo (años)', ticks: { callback: (v) => nf0.format(v) } }),
          y: axis({ type: state.scatterScale, title: 'Monto (miles de millones COP)', beginAtZero: !logScale, ticks: { callback: (v) => (logScale && ![1, 2, 5, 10, 20, 50, 100, 200, 500].includes(v) ? '' : milM(v)) } }),
        },
        plugins: {
          legend: { display: false },
          pointLabels: { items },
          tooltip: tooltipBase({
            title: (it) => it[0]?.raw?.rec ? it[0].raw.rec.ref : '',
            label: (ctx) => { const p = ctx.raw.rec; return p ? `${p.lender}: ${fmtCOP(p.value)}` : ''; },
            afterLabel: (ctx) => { const p = ctx.raw.rec; return p ? [`Plazo: ${fmtYears(p.term)} (${fmtDate(p.start)} → ${fmtDate(p.end)})`, `Carga anual implícita: ${fmtBig(p.annual)}`, `${DIM_LABEL[state.colorBy]}: ${catOf(p)}`] : ''; },
          }),
        },
        onClick: (_e, els, c) => { const el = els[0]; if (!el) return; const p = c.data.datasets[el.datasetIndex].data[el.index].rec; if (p) focusRecord(p); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
      },
    });
    renderLegend('legend-scatter', [...cats.map((c) => ({ label: c, color: colorOfCat(c), shape: 'dot' })), ...(fit ? [{ label: 'Tendencia lineal', shape: 'line', static: true }] : [])], state.hidden.scatter);
    setSpec('scatter', 'Plazo vs. monto por contrato',
      [{ label: 'Contrato' }, { label: 'Acreedor' }, { label: 'Plazo (años)', num: true, fmt: (v) => nf1.format(v) }, { label: 'Monto (COP)', num: true, fmt: fmtCOP }, { label: 'Carga anual (COP)', num: true, fmt: fmtCOP }],
      [...pts].sort((a, b) => a.term - b.term).map((p) => [p.ref, p.lender, p.term, p.value, p.annual]));
  }

  function renderBuckets() {
    const rows = filtered('bucket').filter((r) => Number.isFinite(r.term));
    const total = sum(rows, (r) => r.value);
    const g = BUCKETS.map((b, i) => { const rs = rows.filter((r) => bucketOf(r.term) === i); return { ...b, i, n: rs.length, sum: sum(rs, (r) => r.value), avgTerm: wavg(rs, (r) => r.term) }; });
    const sel = state.filters.bucket;
    upsert('buckets', {
      type: 'bar',
      data: {
        labels: g.map((b) => b.label),
        datasets: [{ label: 'Monto', data: g.map((b) => b.sum / 1e9), backgroundColor: g.map((b) => (sel == null || sel === b.i ? T.s[0] : T.neutral)), keyColor: T.s[0], borderRadius: 4, borderSkipped: 'start', maxBarThickness: 32, categoryPercentage: 0.7 }],
      },
      options: {
        indexAxis: 'y',
        layout: { padding: { right: 70 } },
        scales: { x: axis({ beginAtZero: true, title: 'Miles de millones COP', ticks: { callback: milM, maxTicksLimit: 5 } }), y: axis({ grid: { display: false }, ticks: { color: T.text2 } }) },
        plugins: {
          legend: { display: false },
          barLabels: { formatter: (i) => (g[i].n ? `${g[i].n} · ${pct(total ? g[i].sum / total : 0)}` : null) },
          tooltip: tooltipBase({ label: (ctx) => `Monto: ${fmtCOP(g[ctx.dataIndex].sum)}`, afterLabel: (ctx) => { const b = g[ctx.dataIndex]; return [`${b.n} contratos · ${pct(total ? b.sum / total : 0)} del total`, b.n ? `Monto promedio: ${fmtBig(b.sum / b.n)}` : '']; } }),
        },
        onClick: (_e, els) => { if (!els.length) return; const i = els[0].index; state.filters.bucket = state.filters.bucket === i ? null : i; render(); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
      },
    });
    setSpec('buckets', 'Monto por rango de plazo', [{ label: 'Rango de plazo' }, { label: 'Contratos', num: true }, { label: 'Monto (COP)', num: true, fmt: fmtCOP }, { label: 'Participación', num: true, fmt: pct }],
      g.map((b) => [b.label, b.n, b.sum, total ? b.sum / total : 0]));
  }

  function renderRatio(rows) {
    const pts = rows.filter((r) => Number.isFinite(r.annual)).sort((a, b) => b.annual - a.annual);
    setBoxHeight('ratio', Math.max(220, pts.length * 26 + 60));
    upsert('ratio', {
      type: 'bar',
      data: {
        labels: pts.map((p) => (compact() ? p.labelShort : p.label)),
        datasets: [{ label: 'Carga anual implícita', data: pts.map((p) => p.annual / 1e9), backgroundColor: pts.map((p) => colorOfCat(catOf(p))), keyColor: T.s[0], borderRadius: 4, borderSkipped: 'start', maxBarThickness: 18, categoryPercentage: 0.8 }],
      },
      options: {
        indexAxis: 'y',
        layout: { padding: { right: 44 } },
        scales: { x: axis({ beginAtZero: true, position: 'top', ticks: { callback: milM, maxTicksLimit: 5 } }), y: axis({ grid: { display: false }, ticks: { color: T.text2, autoSkip: false, font: { size: compact() ? 10 : 11 } } }) },
        plugins: {
          legend: { display: false },
          barLabels: { formatter: (i) => nf1.format(pts[i].annual / 1e9) },
          tooltip: tooltipBase({ title: (it) => pts[it[0].dataIndex].ref, label: (ctx) => `${fmtBig(pts[ctx.dataIndex].annual)} por año`, afterLabel: (ctx) => { const p = pts[ctx.dataIndex]; return [`${p.lender} · ${fmtCOP(p.value)}`, `Plazo: ${fmtYears(p.term)}`, `${DIM_LABEL[state.colorBy]}: ${catOf(p)}`]; } }),
        },
        onClick: (_e, els) => { if (els.length) focusRecord(pts[els[0].index]); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
      },
    });
    setSpec('ratio', 'Carga anual implícita por contrato', [{ label: 'Contrato' }, { label: 'Acreedor' }, { label: 'Monto (COP)', num: true, fmt: fmtCOP }, { label: 'Plazo (años)', num: true, fmt: (v) => nf1.format(v) }, { label: 'Carga anual (COP)', num: true, fmt: fmtCOP }],
      pts.map((p) => [p.ref, p.lender, p.value, p.term, p.annual]));
  }

  function renderLenderTerm(rows) {
    const g = groupSum(rows.filter((r) => Number.isFinite(r.term)), (r) => r.lender).map((x) => ({
      ...x, min: Math.min(...x.rows.map((r) => r.term)), max: Math.max(...x.rows.map((r) => r.term)), w: wavg(x.rows, (r) => r.term),
    })).sort((a, b) => b.w - a.w);
    setBoxHeight('lenderTerm', Math.max(200, g.length * 34 + 70));
    upsert('lenderTerm', {
      type: 'bar',
      data: {
        labels: g.map((x) => `${x.key} (${x.n})`),
        datasets: [
          { type: 'bar', label: 'Rango mín.–máx.', data: g.map((x) => [x.min, x.max === x.min ? x.max + 0.04 : x.max]), backgroundColor: alpha(T.s[0], 0.3), keyColor: alpha(T.s[0], 0.5), borderRadius: 4, borderSkipped: false, maxBarThickness: 12 },
          { type: 'line', label: 'Plazo ponderado', data: g.map((x) => x.w), showLine: false, pointRadius: 6, pointHoverRadius: 8, pointBackgroundColor: T.s[0], pointBorderColor: T.surface, pointBorderWidth: 2, keyColor: T.s[0], hitRadius: 12 },
        ],
      },
      options: {
        indexAxis: 'y',
        scales: { x: axis({ title: 'Plazo (años)', min: 0, suggestedMax: 16, ticks: { callback: (v) => nf0.format(v) } }), y: axis({ grid: { display: false }, ticks: { color: T.text2, autoSkip: false } }) },
        plugins: {
          legend: { display: false },
          tooltip: tooltipBase({ title: (it) => g[it[0].dataIndex].key, label: (ctx) => { const x = g[ctx.dataIndex]; return ctx.datasetIndex === 0 ? `Rango: ${nf1.format(x.min)} – ${nf1.format(x.max)} años` : `Ponderado: ${fmtYears(x.w)}`; }, footer: (it) => { const x = g[it[0].dataIndex]; return `${x.n} contratos · ${fmtBig(x.sum)}`; } }),
        },
        onClick: (_e, els) => { if (els.length) toggleSetFilter('lender', g[els[0].index].key); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
      },
    });
    setSpec('lenderTerm', 'Plazo por acreedor', [{ label: 'Acreedor' }, { label: 'Contratos', num: true }, { label: 'Plazo mín.', num: true, fmt: (v) => nf1.format(v) }, { label: 'Plazo máx.', num: true, fmt: (v) => nf1.format(v) }, { label: 'Plazo ponderado', num: true, fmt: (v) => nf1.format(v) }, { label: 'Monto (COP)', num: true, fmt: fmtCOP }],
      g.map((x) => [x.key, x.n, x.min, x.max, x.w, x.sum]));
  }

  function renderLenders() {
    const rows = filtered('lender');
    const total = sum(rows, (r) => r.value);
    const g = groupSum(rows, (r) => r.lender).sort((a, b) => b.sum - a.sum);
    const sel = state.filters.lender;
    setBoxHeight('lenders', Math.max(220, g.length * 34 + 50));
    upsert('lenders', {
      type: 'bar',
      data: { labels: g.map((x) => x.key), datasets: [{ label: 'Monto', data: g.map((x) => x.sum / 1e9), backgroundColor: g.map((x) => (!sel.size || sel.has(x.key) ? T.s[0] : T.neutral)), keyColor: T.s[0], borderRadius: 4, borderSkipped: 'start', maxBarThickness: 22, categoryPercentage: 0.75 }] },
      options: {
        indexAxis: 'y',
        layout: { padding: { right: 60 } },
        scales: { x: axis({ beginAtZero: true, position: 'top', ticks: { callback: milM, maxTicksLimit: 5 } }), y: axis({ grid: { display: false }, ticks: { color: T.text2, autoSkip: false } }) },
        plugins: {
          legend: { display: false },
          barLabels: { formatter: (i) => pct(total ? g[i].sum / total : 0) },
          tooltip: tooltipBase({ label: (ctx) => fmtCOP(g[ctx.dataIndex].sum), afterLabel: (ctx) => { const x = g[ctx.dataIndex]; return [`${x.n} contratos · ${pct(total ? x.sum / total : 0)} del total`, `Plazo ponderado: ${fmtYears(wavg(x.rows, (r) => r.term))}`]; }, footer: () => 'Toque para filtrar' }),
        },
        onClick: (_e, els) => { if (els.length) toggleSetFilter('lender', g[els[0].index].key); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
      },
    });
    setSpec('lenders', 'Monto por acreedor', [{ label: 'Acreedor' }, { label: 'Contratos', num: true }, { label: 'Monto (COP)', num: true, fmt: fmtCOP }, { label: 'Participación', num: true, fmt: pct }, { label: 'Plazo ponderado (años)', num: true, fmt: (v) => nf1.format(v) }],
      g.map((x) => [x.key, x.n, x.sum, total ? x.sum / total : 0, wavg(x.rows, (r) => r.term)]));

    // Concentración (sobre las filas filtradas, incluido el filtro de acreedor)
    const fr = state.rows; const ft = sum(fr, (r) => r.value);
    const shares = groupSum(fr, (r) => r.lender).map((x) => (ft ? x.sum / ft : 0)).sort((a, b) => b - a);
    const hhi = sum(shares, (s) => (s * 100) ** 2);
    const level = hhi < 1500 ? 'baja' : hhi < 2500 ? 'moderada' : 'alta';
    $('#hhi').replaceChildren(
      h('div', null, h('div', { class: 'kpi-label' }, 'Índice HHI', infoBtn('hhi')), h('div', { class: 'kpi-value' }, Number.isFinite(hhi) ? nf0.format(hhi) : '—'), h('div', { class: 'kpi-note' }, `Concentración ${level}`)),
      h('div', null, h('div', { class: 'kpi-label' }, 'Principal acreedor', infoBtn('top1')), h('div', { class: 'kpi-value' }, pct(shares[0] ?? NaN)), h('div', { class: 'kpi-note' }, 'del monto')),
      h('div', null, h('div', { class: 'kpi-label' }, 'Tres principales', infoBtn('top3')), h('div', { class: 'kpi-value' }, pct(sum(shares.slice(0, 3)))), h('div', { class: 'kpi-note' }, 'del monto')),
    );
  }

  function renderLtypeShare() {
    const rows = filtered('ltype');
    const total = sum(rows, (r) => r.value);
    const g = groupSum(rows, (r) => r.ltype);
    const order = [...state.colorMaps.ltype.keys()];
    g.sort((a, b) => order.indexOf(a.key) - order.indexOf(b.key));
    const sel = state.filters.ltype;
    $('#share-ltype').replaceChildren(...g.map((x) => {
      const col = state.colorBy === 'ltype' ? colorOfCat(x.key, 'ltype') : T.s[0];
      const row = h('div', {
        class: 'share-row' + (sel.size && !sel.has(x.key) ? ' is-off' : '') + (sel.has(x.key) ? ' is-sel' : ''), role: 'button', tabindex: '0',
        title: `${x.n} contratos · ${fmtCOP(x.sum)} · toque para filtrar`,
        onclick: () => toggleSetFilter('ltype', x.key),
        onkeydown: (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSetFilter('ltype', x.key); } },
      },
      h('div', { class: 'share-name' }, h('i', { style: `background:${col}` }), x.key),
      h('div', { class: 'share-val' }, h('strong', null, pct(total ? x.sum / total : 0)), ` · ${fmtBig(x.sum)} · ${x.n}`),
      h('div', { class: 'share-track' }, h('div', { class: 'share-fill', style: `width:${total ? (x.sum / total) * 100 : 0}%;background:${col}` })));
      return row;
    }));
    setSpec('ltype', 'Participación por tipo de acreedor', [{ label: 'Tipo' }, { label: 'Contratos', num: true }, { label: 'Monto (COP)', num: true, fmt: fmtCOP }, { label: 'Participación', num: true, fmt: pct }, { label: 'Plazo ponderado (años)', num: true, fmt: (v) => nf1.format(v) }],
      g.map((x) => [x.key, x.n, x.sum, total ? x.sum / total : 0, wavg(x.rows, (r) => r.term)]));
  }

  // Barras apiladas por la dimensión de color: brecha de 2px y esquinas redondeadas solo en el extremo de la pila
  function stackedDatasets(cats, valuesFor, hiddenSet) {
    const visible = cats.filter((c) => !hiddenSet.has(c));
    return cats.map((cat) => {
      const col = colorOfCat(cat);
      return {
        label: cat, data: valuesFor(cat), backgroundColor: col, keyColor: col, borderColor: T.surface,
        borderWidth: { top: 2, bottom: 0, left: 0, right: 0 }, borderSkipped: 'start', maxBarThickness: 28, categoryPercentage: 0.75,
        hidden: hiddenSet.has(cat), stack: 's',
        borderRadius: (ctx) => {
          const i = ctx.dataIndex; const me = visible.indexOf(cat); if (me < 0) return 0;
          const above = visible.slice(me + 1).some((c) => (ctx.chart.data.datasets.find((d) => d.label === c)?.data[i] || 0) > 0);
          return above ? 0 : { topLeft: 4, topRight: 4 };
        },
      };
    });
  }

  function renderYears() {
    const rows = filtered('year').filter((r) => r.signYear);
    if (!rows.length) { upsert('years', { type: 'bar', data: { labels: [], datasets: [] }, options: {} }); return; }
    const years = range(Math.min(...rows.map((r) => r.signYear)), Math.max(...rows.map((r) => r.signYear)));
    const cats = catsOf(rows);
    const f = state.filters;
    const inSel = (y) => (f.yearFrom == null || y >= f.yearFrom) && (f.yearTo == null || y <= f.yearTo);
    const val = (cat) => years.map((y) => sum(rows.filter((r) => r.signYear === y && catOf(r) === cat), (r) => r.value) / 1e9);
    const ds = stackedDatasets(cats, val, state.hidden.years);
    if (f.yearFrom != null || f.yearTo != null) ds.forEach((d) => { const base = d.backgroundColor; d.backgroundColor = years.map((y) => (inSel(y) ? base : alpha(base, 0.25))); });
    const totals = years.map((y) => sum(rows.filter((r) => r.signYear === y && !state.hidden.years.has(catOf(r))), (r) => r.value));
    upsert('years', {
      type: 'bar',
      data: { labels: years.map(String), datasets: ds },
      options: {
        layout: { padding: { top: 18 } },
        scales: { x: axis({ stacked: true, grid: { display: false } }), y: axis({ stacked: true, beginAtZero: true, title: compact() ? '' : 'Miles de millones COP', ticks: { callback: milM, maxTicksLimit: 6 } }) },
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { display: false },
          barLabels: { formatter: (i) => (totals[i] ? nf0.format(totals[i] / 1e9) : null) },
          tooltip: { filter: (it) => !!it.raw, ...tooltipBase({ label: (ctx) => `${ctx.dataset.label}: ${fmtBig(ctx.raw * 1e9)}`, footer: (it) => { const y = years[it[0].dataIndex]; const rs = rows.filter((r) => r.signYear === y); return `Total ${y}: ${fmtBig(sum(rs, (r) => r.value))} · ${rs.length} contratos`; } }) },
        },
        onClick: (_e, els) => { if (!els.length) return; const y = years[els[0].index]; if (f.yearFrom === y && f.yearTo === y) { f.yearFrom = f.yearTo = null; } else { f.yearFrom = f.yearTo = y; } syncControls(); render(); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
      },
    });
    renderLegend('legend-years', cats.map((c) => ({ label: c, color: colorOfCat(c) })), state.hidden.years);
    setSpec('years', 'Monto firmado por año', [{ label: 'Año' }, ...cats.map((c) => ({ label: c, num: true, fmt: fmtCOP })), { label: 'Total (COP)', num: true, fmt: fmtCOP }, { label: 'Contratos', num: true }],
      years.map((y) => [String(y), ...cats.map((c) => sum(rows.filter((r) => r.signYear === y && catOf(r) === c), (r) => r.value)), sum(rows.filter((r) => r.signYear === y), (r) => r.value), rows.filter((r) => r.signYear === y).length]));
  }

  function renderGantt(rows) {
    const pts = rows.filter((r) => r.start && r.end).sort((a, b) => a.start - b.start || b.value - a.value);
    const cats = catsOf(pts);
    const vis = pts.filter((p) => !state.hidden.gantt.has(catOf(p)));
    setBoxHeight('gantt', Math.max(220, vis.length * (compact() ? 22 : 24) + 70));
    const y0 = pts.length ? Math.min(...pts.map((p) => p.start.getFullYear())) : TODAY.getFullYear();
    const y1 = pts.length ? Math.max(...pts.map((p) => p.end.getFullYear())) + 1 : TODAY.getFullYear() + 1;
    upsert('gantt', {
      type: 'bar',
      data: {
        labels: vis.map((p) => (compact() ? p.labelShort : p.label)),
        datasets: [{
          label: 'Vigencia', data: vis.map((p) => [+p.start, +p.end]), backgroundColor: vis.map((p) => colorOfCat(catOf(p))), keyColor: T.s[0],
          borderRadius: 4, borderSkipped: false, maxBarThickness: 14, categoryPercentage: 0.8,
        }],
      },
      options: {
        indexAxis: 'y',
        scales: {
          x: axis({
            type: 'linear', position: 'top', min: +jan1(y0), max: +jan1(y1),
            afterBuildTicks: (sc) => { const step = compact() ? 2 : 1; sc.ticks = range(y0, y1).filter((y) => (y - y0) % step === 0).map((y) => ({ value: +jan1(y) })); },
            ticks: { callback: (v) => new Date(v).getFullYear() },
          }),
          y: axis({ grid: { display: false }, ticks: { color: T.text2, autoSkip: false, font: { size: compact() ? 10 : 11 } } }),
        },
        plugins: {
          legend: { display: false },
          refLine: { value: +TODAY, label: 'Hoy' },
          tooltip: tooltipBase({ title: (it) => vis[it[0].dataIndex].ref, label: (ctx) => { const p = vis[ctx.dataIndex]; return `${fmtDate(p.start)} → ${fmtDate(p.end)}`; }, afterLabel: (ctx) => { const p = vis[ctx.dataIndex]; return [`${p.lender} · ${fmtBig(p.value)}`, `Plazo: ${fmtYears(p.term)} · Remanente: ${fmtYears(p.remaining)}`, `${DIM_LABEL[state.colorBy]}: ${catOf(p)}`]; } }),
        },
        onClick: (_e, els) => { if (els.length) focusRecord(vis[els[0].index]); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
      },
    });
    renderLegend('legend-gantt', [...cats.map((c) => ({ label: c, color: colorOfCat(c) })), { label: 'Hoy', shape: 'vline', static: true }], state.hidden.gantt);
    setSpec('gantt', 'Cronograma de vigencia', [{ label: 'Contrato' }, { label: 'Acreedor' }, { label: 'Inicio', fmt: fmtDate }, { label: 'Fin', fmt: fmtDate }, { label: 'Plazo (años)', num: true, fmt: (v) => nf1.format(v) }, { label: 'Remanente (años)', num: true, fmt: (v) => nf1.format(v) }],
      pts.map((p) => [p.ref, p.lender, p.start, p.end, p.term, p.remaining]));
  }

  function amortIn(r, y) {
    if (!r.start || !r.end || !(r.end > r.start)) return 0;
    const a = Math.max(+r.start, +jan1(y)), b = Math.min(+r.end, +jan1(y + 1));
    return b > a ? r.value * (b - a) / (r.end - r.start) : 0;
  }
  function balanceAt(rows, t) {
    return sum(rows.filter((r) => r.start && r.end && r.start <= t), (r) => r.value * Math.min(1, Math.max(0, (r.end - t) / (r.end - r.start))));
  }

  function renderMaturity(rows) {
    const pts = rows.filter((r) => r.end);
    if (!pts.length) { upsert('maturity', { type: 'bar', data: { labels: [], datasets: [] }, options: {} }); return; }
    const years = range(Math.min(...pts.map((r) => r.end.getFullYear())), Math.max(...pts.map((r) => r.end.getFullYear())));
    const ty = TODAY.getFullYear();
    const vals = years.map((y) => sum(pts.filter((r) => r.end.getFullYear() === y), (r) => r.value));
    upsert('maturity', {
      type: 'bar',
      data: { labels: years.map(String), datasets: [{ label: 'Vence', data: vals.map((v) => v / 1e9), backgroundColor: years.map((y) => (y < ty ? T.neutral : T.s[0])), keyColor: T.s[0], borderRadius: 4, borderSkipped: 'start', maxBarThickness: 28, categoryPercentage: 0.75 }] },
      options: {
        layout: { padding: { top: 18 } },
        scales: { x: axis({ grid: { display: false } }), y: axis({ beginAtZero: true, title: compact() ? '' : 'Miles de millones COP', ticks: { callback: milM, maxTicksLimit: 6 } }) },
        plugins: {
          legend: { display: false },
          refLine: { value: years.indexOf(ty) >= 0 ? years.indexOf(ty) - 0.5 + yearFrac(TODAY) : null, label: 'Hoy' },
          barLabels: { formatter: (i) => (vals[i] ? nf0.format(vals[i] / 1e9) : null) },
          tooltip: tooltipBase({ label: (ctx) => `Vence en ${years[ctx.dataIndex]}: ${fmtBig(vals[ctx.dataIndex])}`, afterLabel: (ctx) => pts.filter((r) => r.end.getFullYear() === years[ctx.dataIndex]).map((r) => `• ${r.lender} ${r.shortRef} (${fmtDate(r.end)})`) }),
        },
      },
    });
    renderLegend('legend-maturity', [{ label: 'Por vencer', color: T.s[0], static: true }, { label: 'Fecha fin ya cumplida', color: T.neutral, static: true }, { label: 'Hoy', shape: 'vline', static: true }]);
    setSpec('maturity', 'Monto por año de vencimiento', [{ label: 'Año' }, { label: 'Contratos', num: true }, { label: 'Monto (COP)', num: true, fmt: fmtCOP }],
      years.map((y, i) => [String(y), pts.filter((r) => r.end.getFullYear() === y).length, vals[i]]));
  }

  function renderAmort(rows) {
    const pts = rows.filter((r) => r.start && r.end);
    if (!pts.length) { upsert('amort', { type: 'bar', data: { labels: [], datasets: [] }, options: {} }); upsert('balance', { type: 'line', data: { labels: [], datasets: [] }, options: {} }); return; }
    const years = range(Math.min(...pts.map((r) => r.start.getFullYear())), Math.max(...pts.map((r) => r.end.getFullYear())));
    const cats = catsOf(pts);
    const ty = TODAY.getFullYear();
    const val = (cat) => years.map((y) => sum(pts.filter((r) => catOf(r) === cat), (r) => amortIn(r, y)) / 1e9);
    const totals = years.map((y) => sum(pts, (r) => amortIn(r, y)));
    upsert('amort', {
      type: 'bar',
      data: { labels: years.map(String), datasets: stackedDatasets(cats, val, state.hidden.amort) },
      options: {
        scales: { x: axis({ stacked: true, grid: { display: false }, ticks: { maxRotation: 0, autoSkipPadding: 6 } }), y: axis({ stacked: true, beginAtZero: true, title: compact() ? '' : 'Miles de millones COP', ticks: { callback: milM, maxTicksLimit: 6 } }) },
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { display: false },
          refLine: { value: years.indexOf(ty) >= 0 ? years.indexOf(ty) - 0.5 + yearFrac(TODAY) : null, label: 'Hoy' },
          tooltip: { filter: (it) => !!it.raw, ...tooltipBase({ label: (ctx) => `${ctx.dataset.label}: ${fmtBig(ctx.raw * 1e9)}`, footer: (it) => `Total ${years[it[0].dataIndex]}: ${fmtBig(totals[it[0].dataIndex])}` }) },
        },
      },
    });
    renderLegend('legend-amort', [...cats.map((c) => ({ label: c, color: colorOfCat(c) })), { label: 'Hoy', shape: 'vline', static: true }], state.hidden.amort);
    setSpec('amort', 'Amortización anual estimada (lineal)', [{ label: 'Año' }, ...cats.map((c) => ({ label: c, num: true, fmt: fmtCOP })), { label: 'Total (COP)', num: true, fmt: fmtCOP }],
      years.map((y, i) => [String(y), ...cats.map((c) => sum(pts.filter((r) => catOf(r) === c), (r) => amortIn(r, y))), totals[i]]));

    // Saldo teórico al cierre de cada año
    const bal = years.map((y) => balanceAt(pts, jan1(y + 1)));
    const idxToday = years.indexOf(ty) >= 0 ? years.indexOf(ty) - 1 + yearFrac(TODAY) : null;
    upsert('balance', {
      type: 'line',
      data: { labels: years.map(String), datasets: [{ label: 'Saldo teórico', data: bal.map((v) => v / 1e9), borderColor: T.s[0], keyColor: T.s[0], backgroundColor: alpha(T.s[0], 0.1), fill: 'origin', borderWidth: 2, pointRadius: 0, pointHoverRadius: 5, pointHoverBackgroundColor: T.s[0], pointHoverBorderColor: T.surface, pointHoverBorderWidth: 2, tension: 0, borderCapStyle: 'round', borderJoinStyle: 'round' }] },
      options: {
        interaction: { mode: 'index', intersect: false },
        scales: { x: axis({ grid: { display: false }, ticks: { maxRotation: 0, autoSkipPadding: 8 } }), y: axis({ beginAtZero: true, title: compact() ? '' : 'Miles de millones COP', ticks: { callback: milM, maxTicksLimit: 6 } }) },
        plugins: {
          legend: { display: false }, crosshair: { enabled: true },
          refLine: { value: idxToday, label: `Hoy: ${fmtBig(balanceAt(pts, TODAY))}` },
          tooltip: tooltipBase({ title: (it) => `Cierre de ${years[it[0].dataIndex]}`, label: (ctx) => `Saldo: ${fmtBig(bal[ctx.dataIndex])}` }),
        },
      },
    });
    setSpec('balance', 'Saldo teórico de capital al cierre de cada año', [{ label: 'Año' }, { label: 'Saldo (COP)', num: true, fmt: fmtCOP }], years.map((y, i) => [String(y), bal[i]]));
  }

  function renderMods(rows) {
    const max = rows.length ? Math.max(0, ...rows.map((r) => r.mods)) : 0;
    const ks = range(0, max);
    const g = ks.map((k) => { const rs = rows.filter((r) => r.mods === k); return { k, n: rs.length, sum: sum(rs, (r) => r.value), rs }; });
    upsert('mods', {
      type: 'bar',
      data: { labels: ks.map((k) => (k === 0 ? 'Ninguna' : String(k))), datasets: [{ label: 'Contratos', data: g.map((x) => x.n), backgroundColor: T.s[0], keyColor: T.s[0], borderRadius: 4, borderSkipped: 'start', maxBarThickness: 28, categoryPercentage: 0.75 }] },
      options: {
        layout: { padding: { top: 18 } },
        scales: { x: axis({ grid: { display: false }, title: 'Número de modificaciones' }), y: axis({ beginAtZero: true, ticks: { precision: 0, maxTicksLimit: 6 }, title: compact() ? '' : 'Contratos' }) },
        plugins: {
          legend: { display: false },
          barLabels: { formatter: (i) => (g[i].n ? String(g[i].n) : null) },
          tooltip: tooltipBase({ label: (ctx) => `${g[ctx.dataIndex].n} contratos · ${fmtBig(g[ctx.dataIndex].sum)}`, afterLabel: (ctx) => g[ctx.dataIndex].rs.slice(0, 8).map((r) => `• ${r.lender} ${r.shortRef}`) }),
        },
      },
    });
    setSpec('mods', 'Contratos según número de modificaciones', [{ label: 'Modificaciones' }, { label: 'Contratos', num: true }, { label: 'Monto (COP)', num: true, fmt: fmtCOP }], g.map((x) => [x.k === 0 ? 'Ninguna' : String(x.k), x.n, x.sum]));
  }

  /* ------------------------------------------------- relación de contratos */
  const COLUMNS = [
    { key: 'ref', label: 'N.º contrato', get: (r) => r.ref, type: 'text', def: true, cls: 'nowrap' },
    { key: 'lender', label: 'Acreedor', get: (r) => r.lender, type: 'text', def: true, cls: 'nowrap' },
    { key: 'value', label: 'Valor', get: (r) => r.value, fmt: fmtCOP, type: 'num', def: true, unit: 1e9, ph: 'ej. >50 (mil M)' },
    { key: 'object', label: 'Objeto', get: (r) => r.object, type: 'long', def: true },
    { key: 'start', label: 'Fecha inicio', get: (r) => r.start, fmt: fmtDate, type: 'date', def: true, ph: 'ej. 2023' },
    { key: 'end', label: 'Fecha fin', get: (r) => r.end, fmt: fmtDate, type: 'date', def: true, ph: 'ej. 2033' },
    { key: 'signed', label: 'Fecha firma', get: (r) => r.signed, fmt: fmtDate, type: 'date', def: true, ph: 'ej. 12/2025' },
    { key: 'term', label: 'Plazo (años)', get: (r) => r.term, fmt: (v) => (Number.isFinite(v) ? nf1.format(v) : '—'), type: 'num', def: true, ph: 'ej. >=10' },
    { key: 'status', label: 'Estado', get: (r) => r.status, type: 'text', def: true, cls: 'nowrap' },
    { key: 'mods', label: 'Modif.', get: (r) => r.mods, fmt: (v) => nf0.format(v), type: 'num', def: false, ph: 'ej. >0' },
    { key: 'remaining', label: 'Remanente (años)', get: (r) => r.remaining, fmt: (v) => (Number.isFinite(v) ? nf1.format(v) : '—'), type: 'num', def: false },
    { key: 'annual', label: 'Carga anual', get: (r) => r.annual, fmt: fmtCOP, type: 'num', def: false, unit: 1e9, ph: 'ej. >10 (mil M)' },
    { key: 'daysAdded', label: 'Días adicionados', get: (r) => r.daysAdded, fmt: (v) => nf0.format(v), type: 'num', def: false },
    { key: 'ltype', label: 'Tipo de acreedor', get: (r) => r.ltype, type: 'text', def: false, cls: 'nowrap' },
    { key: 'period', label: 'Administración', get: (r) => r.period, type: 'text', def: false, cls: 'nowrap' },
    { key: 'nit', label: 'NIT acreedor', get: (r) => r.nit, type: 'text', def: false, cls: 'nowrap' },
    { key: 'secop', label: 'SECOP', get: (r) => r.url, type: 'action', def: true, nosort: true },
  ];
  const colByKey = Object.fromEntries(COLUMNS.map((c) => [c.key, c]));

  function loadColumns() {
    let saved = null;
    try { saved = JSON.parse(store.get(LS_COLS) || 'null'); } catch { saved = null; }
    const order = Array.isArray(saved?.order) ? saved.order.filter((k) => colByKey[k]) : [];
    COLUMNS.forEach((c) => { if (!order.includes(c.key)) order.push(c.key); });
    const visible = Array.isArray(saved?.visible) ? new Set(saved.visible.filter((k) => colByKey[k])) : new Set(COLUMNS.filter((c) => c.def).map((c) => c.key));
    state.columns = { order, visible };
  }
  const saveColumns = () => store.set(LS_COLS, JSON.stringify({ order: state.columns.order, visible: [...state.columns.visible] }));
  const visibleCols = () => state.columns.order.filter((k) => state.columns.visible.has(k)).map((k) => colByKey[k]);

  function cellText(c, r) { const v = c.get(r); return c.fmt ? c.fmt(v) : (v ?? ''); }
  function matchColFilter(c, r, q) {
    if (!q) return true;
    const v = c.get(r);
    if (c.type === 'num') {
      const m = q.replace(/\s/g, '').match(/^(>=|<=|>|<|=)?(-?[\d.,]+)$/);
      if (m) {
        const n = Number(m[2].replace(/\./g, '').replace(',', '.')) * (c.unit || 1);
        if (!Number.isFinite(v)) return false;
        switch (m[1]) { case '>': return v > n; case '<': return v < n; case '>=': return v >= n; case '<=': return v <= n; default: return Math.abs(v - n) < (c.unit ? c.unit / 2 : 0.05); }
      }
    }
    return norm(cellText(c, r)).includes(norm(q));
  }
  function sortedTableRows(rows) {
    const cf = state.colFilters;
    let out = rows.filter((r) => Object.entries(cf).every(([k, q]) => !colByKey[k] || matchColFilter(colByKey[k], r, q)));
    const c = colByKey[state.sort.key];
    if (c && !c.nosort) {
      const dir = state.sort.dir;
      out = [...out].sort((a, b) => {
        let va = c.get(a), vb = c.get(b);
        if (va instanceof Date) va = +va; if (vb instanceof Date) vb = +vb;
        const na = va == null || (typeof va === 'number' && !Number.isFinite(va)), nb = vb == null || (typeof vb === 'number' && !Number.isFinite(vb));
        if (na || nb) return na - nb; // vacíos al final
        return (typeof va === 'number' ? va - vb : String(va).localeCompare(String(vb), 'es', { numeric: true })) * dir;
      });
    }
    return out;
  }

  function secopBtn(r) {
    return r.url ? h('a', { class: 'btn btn-secop', href: r.url, target: '_blank', rel: 'noopener noreferrer', title: `Abrir el proceso ${r.ref} en SECOP II` }, h('span', { class: 'lg' }, 'Ver en '), 'SECOP ↗') : h('span', { class: 'muted' }, 'Sin enlace');
  }

  function renderTable() {
    const cols = visibleCols();
    // Encabezado: etiquetas ordenables y arrastrables + fila de filtros por columna
    const labels = h('tr', { class: 'h-labels' }, cols.map((c) => {
      const active = state.sort.key === c.key;
      const th = h('th', { class: [c.type === 'num' ? 'num' : '', c.key === 'secop' ? 'sticky-r' : ''].join(' ').trim() || null, draggable: 'true', 'data-key': c.key, 'aria-sort': active ? (state.sort.dir > 0 ? 'ascending' : 'descending') : null },
        h('span', { class: 'th-drag', title: 'Arrastre para reordenar', 'aria-hidden': 'true' }, '⋮⋮'),
        c.nosort ? c.label : h('button', { type: 'button', class: 'th-btn', title: 'Ordenar', onclick: () => { state.sort = { key: c.key, dir: active ? -state.sort.dir : (c.type === 'num' || c.type === 'date' ? -1 : 1) }; syncSortMobile(); renderTable(state.rows); } },
          c.label, h('span', { class: 'arrow' }, active ? (state.sort.dir > 0 ? '▲' : '▼') : '')));
      th.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/plain', c.key); e.dataTransfer.effectAllowed = 'move'; });
      th.addEventListener('dragover', (e) => { e.preventDefault(); th.classList.add('drag-over'); });
      th.addEventListener('dragleave', () => th.classList.remove('drag-over'));
      th.addEventListener('drop', (e) => {
        e.preventDefault(); th.classList.remove('drag-over');
        const from = e.dataTransfer.getData('text/plain'); if (!from || from === c.key) return;
        const o = state.columns.order.filter((k) => k !== from); o.splice(o.indexOf(c.key), 0, from);
        state.columns.order = o; saveColumns(); renderColMenu(); renderTable(state.rows);
      });
      return th;
    }));
    const filtersRow = h('tr', { class: 'h-filters' }, cols.map((c) => {
      if (c.type === 'action') return h('th', { class: 'sticky-r' });
      const inp = h('input', { type: 'text', value: state.colFilters[c.key] || '', placeholder: c.ph || 'filtrar…', 'aria-label': `Filtrar columna ${c.label}` });
      inp.addEventListener('input', debounce(() => { state.colFilters[c.key] = inp.value.trim(); if (!state.colFilters[c.key]) delete state.colFilters[c.key]; renderTableBody(); }, 250));
      return h('th', null, inp);
    }));
    $('#tbl-head').replaceChildren(labels, filtersRow);
    renderTableBody();
  }

  function renderTableBody() {
    const cols = visibleCols();
    const tRows = sortedTableRows(state.rows);
    const nFilt = Object.values(state.colFilters).filter(Boolean).length;
    $('#table-count').textContent = `${tRows.length} contrato${tRows.length === 1 ? '' : 's'}${nFilt ? ` (con ${nFilt} filtro${nFilt > 1 ? 's' : ''} de columna)` : ''} · ${fmtBig(sum(tRows, (r) => r.value))}.`;
    const body = tRows.map((r) => h('tr', { 'data-id': r.id }, cols.map((c) => {
      if (c.key === 'secop') return h('td', { class: 'sticky-r' }, secopBtn(r));
      if (c.key === 'object') {
        const t = h('div', { class: 'obj-text', title: 'Toque para ver completo' }, r.object || '—');
        t.addEventListener('click', () => t.classList.toggle('is-open'));
        return h('td', { class: 'obj' }, t);
      }
      if (c.key === 'lender') return h('td', { class: 'nowrap' }, h('span', { class: 'dot', style: `background:${colorOfCat(catOf(r))};margin-right:6px` }), r.lender, r.flags.length ? h('span', { class: 'flag', title: 'Inconsistencias: ' + r.flags.join('; ') }, ' ⚠') : null);
      if (c.key === 'end') return h('td', { class: 'nowrap' + (r.expired ? ' expired' : ''), title: r.expired ? 'Fecha fin ya cumplida' : null }, cellText(c, r));
      return h('td', { class: [c.type === 'num' ? 'num' : '', c.cls || '', c.type === 'date' ? 'nowrap' : ''].join(' ').trim() || null }, cellText(c, r));
    })));
    $('#tbl-body').replaceChildren(...(body.length ? body : [h('tr', null, h('td', { colspan: cols.length, class: 'muted' }, 'Ningún contrato coincide con los filtros.'))]));
    const tot = sum(tRows, (r) => r.value);
    $('#tbl-foot').replaceChildren(h('tr', null, cols.map((c, i) => {
      if (i === 0) return h('td', null, `Total (${tRows.length})`);
      if (c.key === 'value') return h('td', { class: 'num' }, fmtCOP(tot));
      if (c.key === 'term') return h('td', { class: 'num', title: 'Ponderado por monto' }, nf1.format(wavg(tRows, (r) => r.term) || 0));
      if (c.key === 'annual') return h('td', { class: 'num' }, fmtCOP(sum(tRows, (r) => r.annual)));
      if (c.key === 'mods') return h('td', { class: 'num' }, nf0.format(sum(tRows, (r) => r.mods)));
      return h('td', { class: c.key === 'secop' ? 'sticky-r' : null });
    })));
    renderCards(tRows);
  }

  // Vista móvil: tarjetas con la misma información
  function renderCards(tRows) {
    const list = $('#cards-list');
    if (!tRows.length) { list.replaceChildren(h('div', { class: 'c-card muted' }, 'Ningún contrato coincide con los filtros.')); return; }
    list.replaceChildren(...tRows.map((r) => {
      const obj = h('p', { class: 'c-object', title: 'Toque para ver completo' }, r.object || '—');
      obj.addEventListener('click', () => obj.classList.toggle('is-open'));
      return h('article', { class: 'c-card', 'data-id': r.id },
        h('div', { class: 'c-top' },
          h('div', null, h('div', { class: 'c-ref' }, 'Contrato N.º ', r.ref), h('div', { class: 'c-lender' }, h('span', { class: 'dot', style: `background:${colorOfCat(catOf(r))}` }), r.lender, r.flags.length ? h('span', { class: 'flag', title: r.flags.join('; ') }, '⚠') : null)),
          h('span', { class: 'pill' }, r.status)),
        h('div', { class: 'c-value' }, fmtCOP(r.value)),
        h('dl', { class: 'c-grid' },
          h('div', null, h('dt', null, 'Firma'), h('dd', null, fmtDate(r.signed))),
          h('div', null, h('dt', null, 'Inicio'), h('dd', null, fmtDate(r.start))),
          h('div', null, h('dt', null, 'Fin'), h('dd', { class: r.expired ? 'expired' : null }, fmtDate(r.end))),
          h('div', null, h('dt', null, 'Plazo'), h('dd', null, fmtYears(r.term)))),
        obj,
        h('div', { class: 'c-foot' }, h('span', { class: 'muted', style: 'font-size:12px' }, `${r.mods} modificaci${r.mods === 1 ? 'ón' : 'ones'} · ${fmtBig(r.annual)}/año`), secopBtn(r)));
    }));
  }

  function renderColMenu() {
    const panel = $('#col-menu-panel');
    panel.replaceChildren(...state.columns.order.map((k) => {
      const c = colByKey[k];
      const cb = h('input', { type: 'checkbox' }); cb.checked = state.columns.visible.has(k);
      cb.addEventListener('change', () => { cb.checked ? state.columns.visible.add(k) : state.columns.visible.delete(k); if (!state.columns.visible.size) { state.columns.visible.add('ref'); } saveColumns(); renderTable(state.rows); });
      return h('label', null, cb, c.label);
    }));
    $('#col-menu summary').textContent = `Columnas (${state.columns.visible.size}/${COLUMNS.length})`;
  }
  function syncSortMobile() {
    const sel = $('#sort-mobile');
    if (!sel.options.length) {
      COLUMNS.filter((c) => !c.nosort && c.type !== 'long').forEach((c) => {
        sel.append(h('option', { value: `${c.key}:-1` }, `${c.label} ↓`), h('option', { value: `${c.key}:1` }, `${c.label} ↑`));
      });
      sel.addEventListener('change', () => { const [k, d] = sel.value.split(':'); state.sort = { key: k, dir: +d }; renderTable(state.rows); });
    }
    sel.value = `${state.sort.key}:${state.sort.dir}`;
  }

  function focusRecord(r) {
    const desktop = window.matchMedia('(min-width: 900px)').matches || $('#s-contratos').classList.contains('view-table');
    let el = document.querySelector(`${desktop ? '#tbl-body tr' : '#cards-list .c-card'}[data-id="${r.id}"]`);
    if (!el) { state.colFilters = {}; renderTable(state.rows); el = document.querySelector(`${desktop ? '#tbl-body tr' : '#cards-list .c-card'}[data-id="${r.id}"]`); }
    if (!el) return;
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    el.classList.remove('flash'); void el.offsetWidth; el.classList.add('flash');
  }

  /* ------------------------------------------------- ayudas (i) explicativas */
  // Cada indicador y gráfica tiene un botón (i) que explica en palabras sencillas qué muestra, cómo leerlo
  // y, con un ejemplo de «10 manzanas», qué significa el dato con los filtros actuales.
  function metrics() {
    const rows = state.rows;
    const total = sum(rows, (r) => r.value);
    const share = (v) => (total ? v / total : NaN);
    const dated = rows.filter((r) => Number.isFinite(r.term));
    const maxBy = (arr, fn) => arr.reduce((a, r) => (!a || fn(r) > fn(a) ? r : a), null);
    const byLender = groupSum(rows, (r) => r.lender).sort((a, b) => b.sum - a.sum);
    const lenderTerms = byLender.map((g) => ({ name: g.key, w: wavg(g.rows, (r) => r.term) })).filter((x) => Number.isFinite(x.w)).sort((a, b) => b.w - a.w);
    const byYear = groupSum(rows.filter((r) => r.signYear), (r) => r.signYear).sort((a, b) => b.sum - a.sum);
    const byEnd = groupSum(rows.filter((r) => r.end), (r) => r.end.getFullYear()).sort((a, b) => b.sum - a.sum);
    const withDates = rows.filter((r) => r.start && r.end);
    const yrs = withDates.length ? range(Math.min(...withDates.map((r) => r.start.getFullYear())), Math.max(...withDates.map((r) => r.end.getFullYear()))) : [];
    const amortPeak = yrs.map((y) => ({ y, v: sum(withDates, (r) => amortIn(r, y)) })).sort((a, b) => b.v - a.v)[0];
    const balPeak = yrs.map((y) => ({ y, v: balanceAt(withDates, jan1(y + 1)) })).sort((a, b) => b.v - a.v)[0];
    const split = catsOf(rows).map((c) => ({ c, v: sum(rows.filter((r) => catOf(r) === c), (r) => r.value) }));
    const modCounts = countBy(rows, (r) => r.mods);
    const bal = balanceAt(rows, TODAY);
    return {
      n: rows.length, total, share, split,
      wTerm: wavg(rows, (r) => r.term), sTerm: dated.length ? sum(dated, (r) => r.term) / dated.length : NaN,
      wRem: wavg(rows, (r) => r.remaining), bal, balShare: share(bal),
      avg: rows.length ? total / rows.length : NaN, median: median(rows.map((r) => r.value)),
      nMod: rows.filter((r) => r.mods > 0).length, totMods: sum(rows, (r) => r.mods), maxMods: rows.length ? Math.max(...rows.map((r) => r.mods)) : 0,
      noMods: (modCounts.find(([k]) => k === 0) || [0, 0])[1],
      big: maxBy(rows, (r) => r.value), topAnnual: maxBy(dated, (r) => r.annual),
      annualTot: sum(dated, (r) => r.annual),
      b10: share(sum(dated.filter((r) => bucketOf(r.term) === 3), (r) => r.value)),
      byLender, lenderTerms,
      hhi: sum(byLender, (g) => (share(g.sum) * 100) ** 2),
      top3: sum(byLender.slice(0, 3), (g) => share(g.sum)),
      pub: share(sum(rows.filter((r) => r.ltype === TYPE_PUBLIC), (r) => r.value)),
      priv: share(sum(rows.filter((r) => r.ltype === TYPE_PRIVATE), (r) => r.value)),
      yearPeak: byYear[0], endPeak: byEnd[0], amortPeak, balPeak,
      after2027: share(sum(rows.filter((r) => r.end && r.end.getFullYear() > 2027), (r) => r.value)),
      c: state.corr || {},
    };
  }
  const apples = (frac) => (Number.isFinite(frac) ? `${nf1.format(frac * 10)} manzana${Math.abs(frac * 10 - 1) < 0.05 ? '' : 's'}` : '— manzanas');
  const NO_DATA = 'No hay contratos con los filtros actuales; quite algún filtro para ver el ejemplo con datos.';

  const INFO = {
    total: (m) => ({
      t: 'Monto total contratado',
      what: 'Es la suma del valor de todos los contratos de préstamo firmados por la Gobernación (con los filtros que tenga activos). Es lo máximo que los bancos se comprometieron a prestar.',
      how: 'Es la foto completa del endeudamiento contratado. No es lo que se debe hoy: una parte ya se pagó y otra puede no haberse desembolsado todavía. La barra de colores de abajo muestra cómo se reparte según el selector «Colorear gráficas por».',
      ex: m.n ? `Si todo lo que ha pedido prestado el Departamento fueran 10 manzanas, este número son esas 10 manzanas: ${fmtBig(m.total)}. ${m.split.map((x) => `${apples(m.share(x.v))} corresponden a «${x.c}»`).join('; ')}.` : NO_DATA,
    }),
    wterm: (m) => ({
      t: 'Plazo promedio ponderado',
      what: 'Cuántos años tiene en promedio el Departamento para devolver lo que le prestaron. «Ponderado» significa que los préstamos grandes cuentan más que los pequeños.',
      how: 'Más años significa cuotas más pequeñas cada año, pero también que la deuda dura más tiempo y pasa a gobiernos futuros. Si el ponderado es mayor que el promedio simple (donde cada contrato cuenta igual), los préstamos grandes son los de plazo más largo.',
      ex: `Si te prestan 9 manzanas a 10 años y 1 manzana a 2 años, el promedio simple es 6 años, pero el ponderado es 9,2 años, porque casi todas las manzanas están en el préstamo largo. Aquí el ponderado es ${fmtYears(m.wTerm)} y el simple ${fmtYears(m.sTerm)}.`,
    }),
    wrem: (m) => ({
      t: 'Vida remanente ponderada',
      what: 'Cuántos años faltan, en promedio, desde hoy hasta que terminen los contratos. Los préstamos grandes pesan más en el promedio.',
      how: 'Un número alto indica que la deuda seguirá vigente durante muchos años. Los contratos cuya fecha de fin ya pasó cuentan como cero.',
      ex: m.n ? `De 10 manzanas prestadas, en promedio faltan ${fmtYears(m.wRem)} para que se cumpla el plazo de devolverlas.` : NO_DATA,
    }),
    balance: (m) => ({
      t: 'Saldo teórico hoy',
      what: 'Una estimación de cuánto capital faltaría por pagar hoy, suponiendo que cada préstamo se paga en cuotas iguales desde su inicio hasta su fin.',
      how: 'Es aproximado: no conocemos las tablas de pago reales, los periodos de gracia (años en que no se paga capital) ni si todo el dinero se desembolsó. Sirve para dimensionar la deuda, no como cifra oficial.',
      ex: m.n ? `Si te prestan 10 manzanas a 10 años y devuelves 1 cada año, después de 3 años todavía debes 7. Con los contratos actuales, de cada 10 manzanas prestadas aún se deberían unas ${apples(m.balShare)} (${fmtBig(m.bal)}).` : NO_DATA,
    }),
    avg: (m) => ({
      t: 'Monto promedio por contrato',
      what: 'El tamaño típico de un préstamo. El promedio es el total dividido entre el número de contratos. La mediana es el contrato que queda en el medio si se ordenan de menor a mayor.',
      how: 'Si el promedio es bastante mayor que la mediana, hay unos pocos préstamos muy grandes que suben el promedio.',
      ex: `Si tienes tres préstamos de 1, 1 y 10 manzanas, el promedio es 4 manzanas, pero la mediana es 1: la mayoría son pequeños y uno grande jala el promedio hacia arriba. Aquí el promedio es ${fmtBig(m.avg)} y la mediana ${fmtBig(m.median)}.`,
    }),
    mods: (m) => ({
      t: 'Contratos modificados',
      what: 'Cuántos contratos han tenido cambios después de firmados (plazos, condiciones, cláusulas), según lo registrado en SECOP II.',
      how: 'Modificar un contrato es legal y común, pero muchas modificaciones en un mismo contrato merecen revisar por qué se hicieron. En este archivo ninguna modificación aumentó el valor prestado.',
      ex: `Si acuerdas devolver 10 manzanas en 5 años y luego cambias el acuerdo a 10 años, eso es una modificación: debes lo mismo, pero durante más tiempo. Aquí ${m.nMod} de ${m.n} contratos tuvieron cambios, ${nf0.format(m.totMods)} en total.`,
    }),
    biggest: (m) => ({
      t: 'Mayor contrato',
      what: 'El contrato de mayor valor con los filtros actuales.',
      how: 'Muestra cuánto pesa un solo préstamo dentro del total.',
      ex: m.big ? `De cada 10 manzanas prestadas, ${apples(m.share(m.big.value))} vienen solo de este contrato con ${m.big.lender} (${m.big.ref}).` : NO_DATA,
    }),
    pearson: (m) => ({
      t: 'Correlación de Pearson (r)',
      what: 'Un número entre -1 y 1 que indica si los préstamos más grandes tienden a tener plazos más largos (o más cortos).',
      how: 'Cerca de 1: a más plazo, más dinero, casi siempre. Cerca de 0: no hay relación. Cerca de -1: a más plazo, menos dinero. Como guía: menos de 0,2 es muy débil; de 0,2 a 0,4, débil; de 0,4 a 0,6, moderada; más de 0,6, fuerte.',
      ex: Number.isFinite(m.c.r) ? `Si cada vez que pides más manzanas te dan más años para devolverlas, r estaría cerca de 1. Aquí r = ${nf2.format(m.c.r)}, una relación ${strength(m.c.r)}: ${Math.abs(m.c.r) < 0.4 ? 'saber el plazo de un préstamo ayuda poco a adivinar cuántas manzanas se prestaron' : m.c.r > 0 ? 'los préstamos de más manzanas suelen tener más años para devolverse' : 'los préstamos de más manzanas suelen tener menos años para devolverse'}.` : NO_DATA,
    }),
    r2: (m) => ({
      t: 'R² del ajuste lineal',
      what: 'Qué parte de las diferencias de tamaño entre los préstamos se explica por su plazo.',
      how: 'Va de 0 % a 100 %. Un valor bajo significa que el monto depende de otras cosas (el proyecto, el banco, el momento), no del plazo.',
      ex: Number.isFinite(m.c.r2) ? `Si los préstamos tienen tamaños muy distintos, de cada 10 «manzanas de diferencia» entre ellos el plazo explica unas ${apples(m.c.r2)}; el resto se debe a otras razones.` : NO_DATA,
    }),
    spearman: (m) => ({
      t: 'Correlación de Spearman (ρ)',
      what: 'Parecida a Pearson, pero en lugar de usar los montos exactos ordena los préstamos del más pequeño al más grande y del más corto al más largo, y compara los dos órdenes.',
      how: 'Sirve para que un préstamo muy grande no distorsione el resultado. Si Spearman y Pearson dan valores parecidos, la conclusión es más confiable.',
      ex: Number.isFinite(m.c.rho) ? `Pon en una fila tus préstamos de menos a más manzanas y en otra de menos a más años: ρ mide si las dos filas quedan en el mismo orden (1 = idéntico, 0 = sin relación). Aquí ρ = ${nf2.format(m.c.rho)}.` : NO_DATA,
    }),
    slope: (m) => ({
      t: 'Pendiente',
      what: 'Cuánto más grande es, en promedio, un préstamo por cada año adicional de plazo, según la línea punteada de la gráfica.',
      how: 'Es una tendencia promedio, no una regla. Si la correlación es débil, este número es poco confiable.',
      ex: m.c.fit ? `Si la pendiente fuera de 1 manzana por año, un préstamo a 10 años tendría en promedio 1 manzana más que uno a 9 años. Aquí cada año extra de plazo se asocia con ${m.c.fit.b >= 0 ? '' : 'menos '}${fmtBig(Math.abs(m.c.fit.b) * 1e9)}${m.c.fit.b >= 0 ? ' más' : ''} de préstamo.` : NO_DATA,
    }),
    annualTotal: (m) => ({
      t: 'Carga anual implícita total',
      what: 'La suma de lo que habría que devolver de capital cada año si cada préstamo se pagara en cuotas iguales (monto ÷ plazo).',
      how: 'Da una idea del peso anual de la deuda sobre el presupuesto. No incluye intereses, que se pagan aparte y aumentan esta cifra.',
      ex: `Si te prestan 10 manzanas a 10 años, devuelves 1 manzana por año; si te las prestan a 5 años, 2 por año. Sumando todos los contratos, serían unos ${fmtBig(m.annualTot)} de capital por año.`,
    }),
    hhi: (m) => ({
      t: 'Índice HHI (concentración)',
      what: 'Mide si la deuda depende de pocos bancos o está repartida entre muchos. Va de casi 0 a 10.000.',
      how: 'Menos de 1.500: poco concentrada. Entre 1.500 y 2.500: moderada. Más de 2.500: muy concentrada, es decir, pocos bancos tienen mucho peso sobre las finanzas del Departamento.',
      ex: m.n ? `Si un solo vecino te presta las 10 manzanas, el índice es 10.000 (dependes totalmente de él). Si 10 vecinos te prestan 1 manzana cada uno, es 1.000. Aquí es ${nf0.format(m.hhi)}.` : NO_DATA,
    }),
    top1: (m) => ({
      t: 'Principal acreedor',
      what: 'Qué porcentaje de toda la deuda contratada tiene el banco que más le ha prestado al Departamento.',
      how: 'Cuanto más alto, más depende el Departamento de ese banco.',
      ex: m.byLender[0] ? `De cada 10 manzanas, ${apples(m.share(m.byLender[0].sum))} las prestó ${m.byLender[0].key}.` : NO_DATA,
    }),
    top3: (m) => ({
      t: 'Tres principales acreedores',
      what: 'Qué porcentaje de la deuda tienen, sumados, los tres bancos que más han prestado.',
      how: 'Si pasa de la mitad, la mayor parte de la deuda está en pocas manos.',
      ex: m.byLender.length ? `De cada 10 manzanas, ${apples(m.top3)} las prestaron ${m.byLender.slice(0, 3).map((g) => g.key).join(', ')}.` : NO_DATA,
    }),
    scatter: (m) => ({
      t: 'Plazo vs. monto contratado',
      what: 'Cada punto es un contrato. Más a la derecha significa más años para pagar; más arriba, más dinero prestado. La línea punteada es la tendencia promedio. Los colores siguen el selector «Colorear gráficas por».',
      how: 'Si los puntos formaran una escalera que sube hacia la derecha, los préstamos grandes serían siempre los de más plazo. Toque un punto para ubicar el contrato en la tabla. Con «Log» se separan mejor los montos pequeños.',
      ex: m.n ? `Imagine que pide manzanas a varios vecinos: un punto arriba a la derecha es «muchas manzanas y mucho tiempo para devolverlas». Aquí ${apples(m.b10)} de cada 10 se prestaron a cerca de 10 años. Por eso hay una columna de puntos en el 10: el plazo casi siempre fue el mismo, aunque la cantidad prestada cambió mucho.` : NO_DATA,
    }),
    buckets: (m) => ({
      t: 'Monto por rango de plazo',
      what: 'Agrupa los contratos según los años que tienen para pagarse y suma el dinero de cada grupo. Junto a cada barra aparece cuántos contratos hay y qué porcentaje del total representan.',
      how: 'Muestra si la deuda es de corto o de largo plazo. Toque una barra para ver solo esos contratos en todo el tablero.',
      ex: m.n ? `De cada 10 manzanas prestadas, ${apples(m.b10)} se deben devolver en un plazo de entre 9 y 10,5 años.` : NO_DATA,
    }),
    ratio: (m) => ({
      t: 'Carga anual implícita por contrato',
      what: 'Para cada contrato, divide el monto entre los años de plazo: es cuánto capital habría que pagar cada año si se pagara en partes iguales.',
      how: 'Una barra larga es un préstamo que pesa mucho cada año sobre las finanzas, porque es grande o porque su plazo es corto. Permite comparar préstamos de distinto tamaño y plazo.',
      ex: m.topAnnual ? `10 manzanas a 10 años son 1 manzana por año; 10 manzanas a 5 años son 2 por año: el segundo pesa el doble cada año aunque se preste lo mismo. El que más pesa aquí es ${m.topAnnual.lender} (${m.topAnnual.ref}): ${fmtBig(m.topAnnual.value)} a ${fmtYears(m.topAnnual.term)}, unos ${fmtBig(m.topAnnual.annual)} por año.` : NO_DATA,
    }),
    lenderTerm: (m) => ({
      t: 'Plazo por acreedor',
      what: 'Para cada banco, la barra clara va desde el plazo más corto hasta el más largo de sus contratos, y el punto es su plazo promedio (los préstamos grandes pesan más). El número entre paréntesis es cuántos contratos tiene.',
      how: 'Sirve para ver qué bancos prestan a más largo plazo. Si solo se ve el punto, todos los contratos de ese banco tienen el mismo plazo.',
      ex: m.lenderTerms.length > 1 ? `Si un vecino te presta manzanas una vez a 4 años y otra a 10, su barra va de 4 a 10. ${m.lenderTerms[0].name} es el que da más plazo en promedio (${fmtYears(m.lenderTerms[0].w)}) y ${m.lenderTerms[m.lenderTerms.length - 1].name} el que menos (${fmtYears(m.lenderTerms[m.lenderTerms.length - 1].w)}).` : NO_DATA,
    }),
    lenders: (m) => ({
      t: 'Monto contratado por acreedor',
      what: 'Cuánto dinero ha prestado cada banco al Departamento y qué porcentaje del total representa.',
      how: 'Toque una barra para ver solo los contratos de ese banco en todo el tablero; las demás barras se ponen grises. Tóquela otra vez para quitar el filtro.',
      ex: m.byLender.length > 1 ? `De cada 10 manzanas, ${apples(m.share(m.byLender[0].sum))} las prestó ${m.byLender[0].key} y ${apples(m.share(m.byLender[1].sum))} ${m.byLender[1].key}.` : NO_DATA,
    }),
    ltype: (m) => ({
      t: 'Participación por tipo de acreedor',
      what: 'Divide la deuda entre banca pública y de fomento (Findeter, Infivalle y Banco Agrario, entidades del Estado para financiar el desarrollo) y banca privada (bancos comerciales).',
      how: 'La banca de fomento suele ofrecer condiciones especiales (tasas compensadas, plazos largos) y la privada, condiciones de mercado. Toque una fila para filtrar. Debajo están los indicadores de concentración.',
      ex: m.n ? `De cada 10 manzanas, ${apples(m.pub)} vienen de banca pública o de fomento y ${apples(m.priv)} de bancos privados.` : NO_DATA,
    }),
    years: (m) => ({
      t: 'Monto firmado por año',
      what: 'Cuánto dinero se contrató en préstamos cada año, según la fecha de firma. Los colores siguen el selector «Colorear gráficas por» (por defecto, la administración que firmó).',
      how: 'Muestra en qué momentos se aceleró el endeudamiento. Toque una columna para ver solo ese año.',
      ex: m.yearPeak ? `Si cada año pides manzanas, esta gráfica muestra cuántas pediste en cada uno. El año con más fue ${m.yearPeak.key}: ${fmtBig(m.yearPeak.sum)}, es decir, ${apples(m.share(m.yearPeak.sum))} de cada 10 de todo el periodo.` : NO_DATA,
    }),
    gantt: (m) => ({
      t: 'Cronograma de vigencia',
      what: 'Cada barra es un contrato: empieza en su fecha de inicio y termina en su fecha de fin. La línea roja es el día de hoy.',
      how: 'Lo que queda a la derecha de la línea roja es deuda que sigue vigente. Las barras que llegan hasta 2033–2035 son compromisos que tendrán que pagar los próximos gobiernos. Toque una barra para ubicar el contrato en la tabla.',
      ex: m.n ? `Si hoy pides manzanas y las terminas de devolver en 2035, la barra va de hoy a 2035. De cada 10 manzanas, ${apples(m.after2027)} se terminan de devolver después de 2027, cuando ya habrá otro gobierno.` : NO_DATA,
    }),
    maturity: (m) => ({
      t: 'Monto por año de vencimiento',
      what: 'Cuánto dinero de los contratos termina su plazo en cada año. En gris aparecen los contratos cuya fecha de fin ya pasó.',
      how: 'Una barra muy alta indica un año en que terminan muchos créditos a la vez.',
      ex: m.endPeak ? `Si te prestan 10 manzanas en varios préstamos, esta gráfica dice en qué año vence cada uno. En ${m.endPeak.key} vencen contratos por ${fmtBig(m.endPeak.sum)}, es decir, ${apples(m.share(m.endPeak.sum))} de cada 10.` : NO_DATA,
    }),
    amort: (m) => ({
      t: 'Amortización anual estimada',
      what: 'Una estimación de cuánto capital tocaría pagar cada año si cada contrato se pagara en cuotas iguales entre su inicio y su fin.',
      how: 'Es aproximada: los contratos reales pueden tener años de gracia (sin pagar capital) y cuotas distintas, y no incluye intereses. Sirve para ver en qué años se acumula la carga.',
      ex: m.amortPeak ? `Si te prestan 10 manzanas a 10 años, devuelves 1 cada año. Si dos años después pides otras 10 a 10 años, durante varios años devuelves 2 por año. Aquí el año de mayor carga estimada es ${m.amortPeak.y}, con ${fmtBig(m.amortPeak.v)}.` : NO_DATA,
    }),
    balanceChart: (m) => ({
      t: 'Saldo teórico de capital',
      what: 'Cuánto capital quedaría por pagar al final de cada año, sumando todos los contratos, si se pagaran en cuotas iguales.',
      how: 'La curva sube cuando se firman préstamos nuevos y baja a medida que se pagan. El punto más alto es cuando la deuda estimada fue mayor. La línea roja marca hoy.',
      ex: m.balPeak ? `Si te prestan 10 manzanas a 10 años, al final del primer año debes 9, al final del segundo 8, y así hasta llegar a 0. Hoy la deuda estimada es ${fmtBig(m.bal)}: de cada 10 manzanas prestadas aún se deberían ${apples(m.balShare)}. El punto más alto es al cierre de ${m.balPeak.y}, con ${fmtBig(m.balPeak.v)}.` : NO_DATA,
    }),
    modsChart: (m) => ({
      t: 'Contratos según número de modificaciones',
      what: 'Cuántos contratos no han tenido cambios y cuántos han tenido 1, 2, 3 o más modificaciones registradas en SECOP II.',
      how: 'Muchas modificaciones en un mismo contrato pueden indicar renegociaciones frecuentes; vale la pena revisar en SECOP qué se cambió en cada una.',
      ex: m.n ? `Si acuerdas devolver 10 manzanas y después cambias el acuerdo tres veces, ese préstamo tiene 3 modificaciones. Aquí ${m.noMods} contratos nunca se han modificado y el que más cambios tiene acumula ${m.maxMods}.` : NO_DATA,
    }),
  };

  function infoBtn(key) {
    if (!INFO[key]) return null;
    return h('button', { type: 'button', class: 'info-btn', 'aria-label': 'Qué significa este dato', title: '¿Qué significa?', 'data-info': key }, 'i');
  }
  function openInfo(key) {
    const def = INFO[key]; if (!def) return;
    const d = def(metrics());
    let dlg = document.getElementById('info-dialog');
    if (!dlg) {
      dlg = h('dialog', { id: 'info-dialog', class: 'info-dialog', 'aria-labelledby': 'info-title' });
      dlg.addEventListener('click', (e) => { if (e.target === dlg) dlg.close(); });
      document.body.append(dlg);
    }
    dlg.replaceChildren(h('div', { class: 'info-body' },
      h('div', { class: 'info-head' }, h('h3', { id: 'info-title' }, d.t), h('button', { type: 'button', class: 'info-close', 'aria-label': 'Cerrar', onclick: () => dlg.close() }, '×')),
      h('p', { class: 'info-label' }, 'Qué muestra'), h('p', null, d.what),
      h('p', { class: 'info-label' }, 'Cómo leerlo'), h('p', null, d.how),
      d.ex ? h('div', { class: 'info-example' }, h('p', { class: 'info-label' }, '🍎 Ejemplo con manzanas'), h('p', null, d.ex)) : null,
      state.rows.length !== state.all.length ? h('p', { class: 'info-foot' }, 'Las cifras del ejemplo usan los filtros activos.') : null));
    if (typeof dlg.showModal === 'function') dlg.showModal(); else dlg.setAttribute('open', '');
  }
  function addChartInfoButtons() {
    const map = { balance: 'balanceChart', mods: 'modsChart' };
    document.querySelectorAll('[data-chart]').forEach((card) => {
      const h3 = card.querySelector('.card-head h3');
      if (!h3 || h3.querySelector('.info-btn')) return;
      h3.append(' ', infoBtn(map[card.dataset.chart] || card.dataset.chart));
    });
  }

  /* ------------------------------------------------------------ metodología */
  function renderMethodology() {
    const all = state.all;
    const flagged = all.filter((r) => r.flags.length);
    const refCounts = countBy(all.filter((r) => r.ref !== '—'), (r) => r.ref).filter(([, n]) => n > 1);
    const expired = all.filter((r) => r.expired && !r.flags.length);
    const added = all.filter((r) => r.daysAdded > 0);
    const li = (t) => h('li', null, t);
    $('#methodology').replaceChildren(...[
      h('h3', null, 'Universo de datos'),
      h('p', null, `Se analizan ${all.length} contratos de la pestaña «${state.sheetName}» del archivo ${state.fileName}. ${flagged.length ? `${flagged.length} registro(s) presentan inconsistencias y se excluyen por defecto (active «Incluir registros con inconsistencias» para verlos).` : ''}`),
      flagged.length ? h('ul', null, flagged.map((r) => li(`${r.ref} · ${r.lenderRaw} · ${fmtCOP(r.value)}: ${r.flags.join('; ')}.`))) : null,
      refCounts.length ? h('p', { style: 'margin-top:8px' }, 'Referencias de contrato repetidas en registros distintos (verificar en SECOP):') : null,
      refCounts.length ? h('ul', null, refCounts.map(([ref, n]) => li(`${ref}: ${n} registros (${all.filter((r) => r.ref === ref).map((r) => `${r.lender}, firmado ${fmtDate(r.signed)}`).join(' / ')})`))) : null,
      expired.length ? h('p', { style: 'margin-top:8px' }, 'Contratos cuya fecha de fin ya se cumplió pero su estado en SECOP no es «terminado/liquidado»:') : null,
      expired.length ? h('ul', null, expired.map((r) => li(`${r.ref} · ${r.lender} · fin ${fmtDate(r.end)} · estado «${r.status}»`))) : null,
      added.length ? h('p', { style: 'margin-top:8px' }, 'Contratos con días adicionados por modificación:') : null,
      added.length ? h('ul', null, added.map((r) => li(`${r.ref} · ${r.lender}: ${nf0.format(r.daysAdded)} días (≈ ${nf1.format(r.daysAdded / YEAR_DAYS)} años)`))) : null,
      h('h3', null, 'Definiciones'),
      h('ul', null,
        li('Plazo: años entre la fecha de inicio y la fecha de fin del contrato (días ÷ 365,25).'),
        li('Plazo promedio ponderado: promedio de plazos ponderado por el valor de cada contrato.'),
        li('Vida remanente: años entre hoy y la fecha de fin (0 si ya se cumplió).'),
        li('Carga anual implícita: valor ÷ plazo. Indica cuánto capital habría que amortizar por año con pagos lineales; permite comparar contratos de distinto plazo.'),
        li('Amortización y saldo teórico: distribución lineal del valor entre inicio y fin, sin periodos de gracia, desembolsos parciales ni intereses. Es una aproximación para dimensionar el perfil de la deuda, no el saldo real.'),
        li('Correlación de Pearson (r) mide relación lineal; Spearman (ρ) usa rangos y es menos sensible a valores extremos. R² = r². Con pocos contratos, interpretar con cautela.'),
        li('Índice Herfindahl-Hirschman (HHI): suma de cuadrados de las participaciones (%) por acreedor. < 1.500 baja, 1.500–2.500 moderada, > 2.500 alta concentración.'),
        li('Administración: periodo de gobierno de cuatro años según la fecha de firma (2020–2023, 2024–2027).'),
        li('Tipo de acreedor: «Banca pública y de fomento» agrupa Findeter, Infivalle y Banco Agrario; el resto se clasifica como «Banca privada».'),
      ),
    ].filter(Boolean));
  }

  /* ------------------------------------------------------- controles/filtros */
  const MS_DIMS = ['lender', 'ltype', 'period', 'status'];
  function buildMultiSelects() {
    MS_DIMS.forEach((dim) => {
      const host = document.querySelector(`[data-ms="${dim}"]`);
      host.querySelector('details')?.remove();
      const opts = countBy(state.all, (r) => r[dim]).sort((a, b) => (dim === 'period' ? a[0].localeCompare(b[0]) : b[1] - a[1]));
      const summary = h('summary', null, 'Todos');
      const panel = h('div', { class: 'ms-panel' },
        opts.map(([v, n]) => {
          const cb = h('input', { type: 'checkbox', value: v });
          cb.addEventListener('change', () => { cb.checked ? state.filters[dim].add(v) : state.filters[dim].delete(v); render(); });
          return h('label', null, cb, h('span', null, v), h('span', { class: 'ms-count' }, n));
        }),
        h('div', { class: 'ms-actions' }, h('button', { type: 'button', class: 'btn btn-ghost btn-sm', onclick: () => { state.filters[dim].clear(); syncControls(); render(); } }, 'Todos')));
      const det = h('details', { class: 'ms', 'data-dim': dim }, summary, panel);
      det.addEventListener('toggle', () => { if (det.open) document.querySelectorAll('.filter-panel details.ms').forEach((d) => { if (d !== det) d.open = false; }); });
      host.append(det);
    });
    const years = [...new Set(state.all.map((r) => r.signYear).filter(Boolean))].sort();
    for (const id of ['#f-year-from', '#f-year-to']) {
      $(id).replaceChildren(h('option', { value: '' }, 'Todos'), ...years.map((y) => h('option', { value: y }, y)));
    }
    const nf = state.all.filter((r) => r.flags.length).length;
    $('#flagged-count').textContent = nf ? `(${nf})` : '';
    $('#f-include-flagged').closest('label').hidden = !nf;
  }

  function syncControls() {
    const f = state.filters;
    $('#f-year-from').value = f.yearFrom ?? '';
    $('#f-year-to').value = f.yearTo ?? '';
    $('#f-term-min').value = f.termMin ?? ''; $('#f-term-max').value = f.termMax ?? '';
    $('#f-amt-min').value = f.amtMin ?? ''; $('#f-amt-max').value = f.amtMax ?? '';
    $('#f-search').value = f.q; $('#f-include-flagged').checked = f.includeFlagged; $('#f-colorby').value = state.colorBy;
    MS_DIMS.forEach((dim) => {
      const det = document.querySelector(`details.ms[data-dim="${dim}"]`); if (!det) return;
      det.querySelectorAll('input[type=checkbox]').forEach((cb) => { cb.checked = f[dim].has(cb.value); });
      const s = f[dim];
      det.querySelector('summary').textContent = !s.size ? 'Todos' : s.size === 1 ? [...s][0] : `${s.size} seleccionados`;
    });
  }

  function toggleSetFilter(dim, v) {
    const s = state.filters[dim];
    if (s.size === 1 && s.has(v)) s.clear(); else { s.clear(); s.add(v); }
    syncControls(); render();
  }

  function activeChips() {
    const f = state.filters; const chips = [];
    if (f.yearFrom != null || f.yearTo != null) chips.push({ text: `Firma: ${f.yearFrom ?? '…'}–${f.yearTo ?? '…'}`, clear: () => { f.yearFrom = f.yearTo = null; } });
    MS_DIMS.forEach((dim) => { if (f[dim].size) chips.push({ text: `${DIM_LABEL[dim]}: ${[...f[dim]].join(', ')}`, clear: () => f[dim].clear() }); });
    if (f.termMin != null || f.termMax != null) chips.push({ text: `Plazo: ${f.termMin ?? 0}–${f.termMax ?? '∞'} años`, clear: () => { f.termMin = f.termMax = null; } });
    if (f.amtMin != null || f.amtMax != null) chips.push({ text: `Monto: ${f.amtMin ?? 0}–${f.amtMax ?? '∞'} mil M`, clear: () => { f.amtMin = f.amtMax = null; } });
    if (f.bucket != null) chips.push({ text: `Rango: ${BUCKETS[f.bucket].label}`, clear: () => { f.bucket = null; } });
    if (f.q) chips.push({ text: `Buscar: «${f.q}»`, clear: () => { f.q = ''; } });
    if (f.includeFlagged && state.all.some((r) => r.flags.length)) chips.push({ text: 'Incluye inconsistencias', clear: () => { f.includeFlagged = false; } });
    return chips;
  }
  function renderChips() {
    const chips = activeChips();
    $('#chips').replaceChildren(...chips.map((c) => h('span', { class: 'chip' }, c.text, h('button', { type: 'button', 'aria-label': `Quitar filtro ${c.text}`, onclick: () => { c.clear(); syncControls(); render(); } }, '×'))));
    const b = $('#filter-badge'); b.hidden = !chips.length; b.textContent = chips.length;
  }

  function wireControls() {
    const numOrNull = (v) => (v === '' || v == null || !Number.isFinite(+v) ? null : +v);
    const f = () => state.filters;
    const on = (id, ev, fn) => $(id).addEventListener(ev, fn);
    on('#f-year-from', 'change', (e) => { f().yearFrom = numOrNull(e.target.value); render(); });
    on('#f-year-to', 'change', (e) => { f().yearTo = numOrNull(e.target.value); render(); });
    const deb = debounce(() => render(), 250);
    on('#f-term-min', 'input', (e) => { f().termMin = numOrNull(e.target.value); deb(); });
    on('#f-term-max', 'input', (e) => { f().termMax = numOrNull(e.target.value); deb(); });
    on('#f-amt-min', 'input', (e) => { f().amtMin = numOrNull(e.target.value); deb(); });
    on('#f-amt-max', 'input', (e) => { f().amtMax = numOrNull(e.target.value); deb(); });
    on('#f-search', 'input', (e) => { f().q = e.target.value.trim(); deb(); });
    on('#f-include-flagged', 'change', (e) => { f().includeFlagged = e.target.checked; render(); });
    on('#f-colorby', 'change', (e) => { state.colorBy = e.target.value; Object.values(state.hidden).forEach((s) => s.clear()); render(); });
    on('#btn-reset', 'click', () => { state.filters = blankFilters(); state.colFilters = {}; syncControls(); render(); renderTable(state.rows); });
    const filtersEl = document.querySelector('.filters');
    const setOpen = (open) => { filtersEl.classList.toggle('is-open', open); $('#btn-filters').setAttribute('aria-expanded', String(open)); };
    on('#btn-filters', 'click', () => setOpen(!filtersEl.classList.contains('is-open')));
    on('#btn-apply', 'click', () => setOpen(false));
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') setOpen(false); });
    document.addEventListener('click', (e) => {
      if (!filtersEl.contains(e.target)) setOpen(false);
      const cm = $('#col-menu'); if (cm.open && !cm.contains(e.target)) cm.open = false;
    });
    document.querySelectorAll('.seg-btn[data-scale]').forEach((b) => b.addEventListener('click', () => {
      state.scatterScale = b.dataset.scale;
      document.querySelectorAll('.seg-btn[data-scale]').forEach((x) => x.classList.toggle('is-on', x === b));
      render('none');
    }));
    document.querySelectorAll('.seg-btn[data-view]').forEach((b) => b.addEventListener('click', () => {
      $('#s-contratos').classList.toggle('view-table', b.dataset.view === 'table');
      document.querySelectorAll('.seg-btn[data-view]').forEach((x) => x.classList.toggle('is-on', x === b));
    }));
    on('#btn-col-reset', 'click', () => { store.set(LS_COLS, ''); loadColumns(); renderColMenu(); renderTable(state.rows); });
    on('#btn-theme', 'click', () => {
      const next = isDark() ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next); store.set(LS_THEME, next); render('none');
    });
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => render('none'));
    let lastCompact = compact();
    window.addEventListener('resize', debounce(() => { if (compact() !== lastCompact) { lastCompact = compact(); render('none'); } }, 200));
    on('#file-input', 'change', (e) => {
      const file = e.target.files && e.target.files[0]; if (!file) return;
      file.arrayBuffer().then((buf) => loadWorkbook(buf, file.name)).catch((err) => showNotice(`No se pudo leer el archivo: ${err.message}`));
      e.target.value = '';
    });
    on('#btn-xlsx', 'click', exportExcel);
    on('#btn-pdf', 'click', () => exportPDF().catch((err) => { $('#busy').hidden = true; alert('No se pudo generar el PDF: ' + err.message); }));
  }

  /* -------------------------------------------------------------- URL hash */
  function writeHash() {
    const f = state.filters; const p = new URLSearchParams();
    if (f.yearFrom != null) p.set('y1', f.yearFrom); if (f.yearTo != null) p.set('y2', f.yearTo);
    MS_DIMS.forEach((d) => { if (f[d].size) p.set(d, [...f[d]].join('|')); });
    if (f.termMin != null) p.set('tmin', f.termMin); if (f.termMax != null) p.set('tmax', f.termMax);
    if (f.amtMin != null) p.set('amin', f.amtMin); if (f.amtMax != null) p.set('amax', f.amtMax);
    if (f.bucket != null) p.set('b', f.bucket); if (f.q) p.set('q', f.q); if (f.includeFlagged) p.set('inc', '1');
    if (state.colorBy !== 'period') p.set('c', state.colorBy);
    const s = p.toString();
    history.replaceState(null, '', s ? '#' + s : location.pathname + location.search);
  }
  function readHash() {
    const p = new URLSearchParams(location.hash.slice(1)); const f = state.filters;
    const n = (k) => (p.has(k) && Number.isFinite(+p.get(k)) ? +p.get(k) : null);
    f.yearFrom = n('y1'); f.yearTo = n('y2'); f.termMin = n('tmin'); f.termMax = n('tmax'); f.amtMin = n('amin'); f.amtMax = n('amax');
    f.bucket = n('b'); if (f.bucket != null && !BUCKETS[f.bucket]) f.bucket = null;
    MS_DIMS.forEach((d) => { f[d] = new Set(p.has(d) ? p.get(d).split('|').filter(Boolean) : []); });
    f.q = p.get('q') || ''; f.includeFlagged = p.get('inc') === '1';
    if (['period', 'ltype', 'status'].includes(p.get('c'))) state.colorBy = p.get('c');
  }

  /* ----------------------------------------------------------------- render */
  function render(mode) {
    if (!state.all.length) return;
    readTokens();
    const rows = filtered();
    state.rows = rows;
    renderChips();
    $('#result-count').textContent = `${rows.length}/${state.all.filter((r) => state.filters.includeFlagged || !r.flags.length).length}`;
    renderKPIs(rows);
    renderCorrelation(rows);
    renderBuckets();
    renderRatio(rows);
    renderLenderTerm(rows);
    renderLenders();
    renderLtypeShare();
    renderYears();
    renderGantt(rows);
    renderMaturity(rows);
    renderAmort(rows);
    renderMods(rows);
    renderTable(rows);
    renderSpecToggles();
    Object.values(charts).forEach((c) => { if (mode === 'none') c.update('none'); });
    writeHash();
  }

  function showNotice(msg, withUpload) {
    const n = $('#load-notice'); n.hidden = false;
    n.replaceChildren(h('div', null, h('p', null, msg), withUpload ? h('p', { class: 'muted', style: 'margin-top:6px' }, 'Use el botón «Cargar Excel» para seleccionar el archivo.') : null));
  }

  function loadWorkbook(buf, name) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const sheetName = wb.SheetNames.find((s) => norm(s) === norm(SHEET_NAME)) || wb.SheetNames[0];
    const { records: recs } = parseSheet(wb.Sheets[sheetName]);
    if (!recs.length) throw new Error(`la pestaña «${sheetName}» no tiene contratos`);
    state.all = recs; state.fileName = name; state.sheetName = sheetName;
    $('#load-notice').hidden = true;
    $('#source-line').textContent = `Fuente: SECOP II · ${name} · pestaña «${sheetName}» · ${recs.length} contratos`;
    $('#footer-file').textContent = name;
    readTokens();
    buildColorMaps();
    buildMultiSelects();
    readHash();
    syncControls();
    renderMethodology();
    render();
  }

  /* ---------------------------------------------------------- exportaciones */
  function filtersSummary() {
    const chips = activeChips().map((c) => c.text);
    const cf = Object.entries(state.colFilters).filter(([, v]) => v).map(([k, v]) => `${colByKey[k].label} «${v}»`);
    if (cf.length) chips.push('Filtros de columna en la tabla: ' + cf.join(', '));
    return chips.length ? chips : ['Sin filtros (todos los contratos válidos)'];
  }
  const stamp = () => { const d = new Date(); return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}${String(d.getMinutes()).padStart(2, '0')}`; };
  const nowText = () => new Date().toLocaleString('es-CO', { dateStyle: 'long', timeStyle: 'short' });
  const SHEET_ORDER = ['lenders', 'ltype', 'buckets', 'lenderTerm', 'ratio', 'years', 'maturity', 'amort', 'balance', 'mods'];

  function exportExcel() {
    if (!state.all.length) return;
    const wb = XLSX.utils.book_new();
    const tRows = sortedTableRows(state.rows);
    const c = state.corr || {};
    const resumen = [
      ['Contratos de crédito público · Gobernación del Valle del Cauca (Hacienda)'],
      [`Reporte generado: ${nowText()}`], [`Fuente: ${state.fileName} · pestaña «${state.sheetName}»`], [],
      ['Filtros aplicados'], ...filtersSummary().map((t) => [t]), [],
      ['Indicador', 'Valor', 'Detalle'], ...state.kpis.map((k) => [k.label, Number.isFinite(k.raw) ? (Math.abs(k.raw) >= 1e6 ? Math.round(k.raw) : +k.raw.toFixed(2)) : k.value, k.note]), [],
      ['Relación plazo – monto', 'Valor'],
      ['Contratos analizados', c.n], ['Correlación de Pearson (r)', Number.isFinite(c.r) ? +c.r.toFixed(4) : null], ['R²', Number.isFinite(c.r2) ? +c.r2.toFixed(4) : null],
      ['Correlación de Spearman (ρ)', Number.isFinite(c.rho) ? +c.rho.toFixed(4) : null], ['Pendiente (mil M COP por año de plazo)', c.fit ? +c.fit.b.toFixed(3) : null], ['Intercepto (mil M COP)', c.fit ? +c.fit.a.toFixed(3) : null], [],
      ['Hallazgos'], ...state.insight.map((t) => [t]),
    ];
    const ws0 = XLSX.utils.aoa_to_sheet(resumen);
    Object.keys(ws0).forEach((a) => { if (a[0] !== '!' && ws0[a].t === 'n' && Math.abs(ws0[a].v) >= 1e6) ws0[a].z = '#,##0'; });
    ws0['!cols'] = [{ wch: 48 }, { wch: 24 }, { wch: 60 }];
    XLSX.utils.book_append_sheet(wb, ws0, 'Resumen');

    const head = ['N.º contrato', 'Acreedor', 'Acreedor (razón social)', 'NIT', 'Tipo de acreedor', 'Valor (COP)', 'Objeto', 'Fecha firma', 'Fecha inicio', 'Fecha fin', 'Plazo (años)', 'Vida remanente (años)', 'Carga anual implícita (COP)', 'Estado', 'N.º modificaciones', 'Tipos de modificación', 'Días adicionados', 'Administración', 'Inconsistencias', 'URL SECOP'];
    const data = tRows.map((r) => [r.ref, r.lender, r.lenderRaw, r.nit, r.ltype, r.value, r.object, r.signed, r.start, r.end,
      Number.isFinite(r.term) ? +r.term.toFixed(2) : null, Number.isFinite(r.remaining) ? +r.remaining.toFixed(2) : null, Number.isFinite(r.annual) ? Math.round(r.annual) : null,
      r.status, r.mods, r.modTypes, r.daysAdded, r.period, r.flags.join('; '), r.url]);
    const ws1 = XLSX.utils.aoa_to_sheet([head, ...data], { cellDates: true, dateNF: 'dd/mm/yyyy' });
    for (let i = 0; i < data.length; i++) {
      [5, 12].forEach((col) => { const a = XLSX.utils.encode_cell({ r: i + 1, c: col }); if (ws1[a]) ws1[a].z = '#,##0'; });
      [7, 8, 9].forEach((col) => { const a = XLSX.utils.encode_cell({ r: i + 1, c: col }); if (ws1[a]) ws1[a].z = 'dd/mm/yyyy'; });
      const u = XLSX.utils.encode_cell({ r: i + 1, c: 19 }); if (ws1[u] && tRows[i].url) ws1[u].l = { Target: tRows[i].url };
    }
    ws1['!cols'] = [18, 20, 40, 12, 24, 18, 60, 12, 12, 12, 10, 10, 18, 14, 10, 22, 10, 24, 30, 50].map((w) => ({ wch: w }));
    ws1['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.length, c: head.length - 1 } }) };
    XLSX.utils.book_append_sheet(wb, ws1, 'Contratos');

    for (const id of SHEET_ORDER) {
      const s = state.specs[id]; if (!s) continue;
      const rows = s.rows.map((r) => r.map((v, i) => (v instanceof Date ? v : typeof v === 'number' ? (Number.isFinite(v) ? (s.columns[i].fmt === fmtCOP ? Math.round(v) : +v.toFixed(4)) : null) : v)));
      const ws = XLSX.utils.aoa_to_sheet([s.columns.map((c) => c.label), ...rows], { cellDates: true });
      s.columns.forEach((col, ci) => { if (col.fmt === fmtCOP) for (let ri = 0; ri < rows.length; ri++) { const a = XLSX.utils.encode_cell({ r: ri + 1, c: ci }); if (ws[a]) ws[a].z = '#,##0'; } if (col.fmt === pct) for (let ri = 0; ri < rows.length; ri++) { const a = XLSX.utils.encode_cell({ r: ri + 1, c: ci }); if (ws[a]) ws[a].z = '0.0%'; } });
      ws['!cols'] = s.columns.map((col, i) => ({ wch: i === 0 ? 30 : 18 }));
      XLSX.utils.book_append_sheet(wb, ws, s.title.replace(/[\\/?*[\]:]/g, '').slice(0, 31));
    }
    XLSX.writeFile(wb, `reporte_credito_publico_valle_${stamp()}.xlsx`);
  }

  // jsPDF (fuentes estándar) solo admite WinAnsi: se sustituyen símbolos fuera de ese juego.
  const pdfText = (s) => String(s ?? '').replace(/[  ]/g, ' ').replace(/≈/g, '~').replace(/≥/g, '>=').replace(/≤/g, '<=').replace(/⇒/g, '=>').replace(/→/g, '->').replace(/↗/g, '').replace(/⚠/g, '!').replace(/ρ/g, 'rho').replace(/²/g, '2').replace(/∞/g, 'inf').replace(/[–—]/g, '-').replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/•/g, '-').replace(/…/g, '...').replace(/[^\x00-\xff]/g, '');

  const hexRgb = (hex) => { const m = hex.replace('#', ''); const n = parseInt(m, 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
  async function exportPDF() {
    if (!state.all.length) return;
    $('#busy').hidden = false;
    await new Promise((r) => setTimeout(r, 30));
    // Las imágenes del PDF siempre se generan en tema claro
    const prevTheme = document.documentElement.getAttribute('data-theme');
    document.documentElement.setAttribute('data-theme', 'light');
    render('none');
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    // Cada gráfica se captura a un ancho fijo para que la escala del texto sea homogénea en el PDF
    const imgs = {};
    const legendCats = catsOf(state.rows).map((c) => [c, colorOfCat(c)]);
    const LEGEND_IDS = ['scatter', 'ratio', 'years', 'gantt', 'amort'];
    for (const [id, c] of Object.entries(charts)) {
      const box = c.canvas.parentNode; const prev = box.style.cssText;
      const hpx = box.style.height ? parseFloat(box.style.height) : 340;
      box.style.width = '680px'; box.style.height = hpx + 'px';
      c.stop(); c.resize(680, hpx); c.update('none');
      imgs[id] = { src: c.toBase64Image('image/png', 1), w: c.width, h: c.height, legend: LEGEND_IDS.includes(id) ? legendCats : id === 'maturity' ? [['Por vencer', T.s[0]], ['Fecha fin ya cumplida', T.neutral]] : null };
      box.style.cssText = prev; c.stop(); c.resize();
    }
    if (prevTheme) document.documentElement.setAttribute('data-theme', prevTheme); else document.documentElement.removeAttribute('data-theme');
    render('none');

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight(), M = 12;
    const blue = [42, 120, 214], ink = [11, 11, 11], ink2 = [82, 81, 78];
    let y = M;
    const header = () => {
      doc.setFillColor(...blue); doc.rect(0, 0, W, 3, 'F');
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...ink2);
      doc.text(pdfText('Contratos de crédito público · Gobernación del Valle del Cauca'), M, 8);
    };
    const newPage = () => { doc.addPage(); header(); y = 14; };
    const ensure = (need) => { if (y + need > H - M) newPage(); };

    header();
    doc.setFont('helvetica', 'bold'); doc.setFontSize(18); doc.setTextColor(...ink);
    doc.text(pdfText('Reporte de contratos de crédito público 2020–2026'), M, 18);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(...ink2);
    doc.text(pdfText(`Gobernación del Valle del Cauca · Secretaría de Hacienda · Generado: ${nowText()}`), M, 24);
    doc.text(pdfText(`Fuente: SECOP II · ${state.fileName} · pestaña «${state.sheetName}»`), M, 29);
    y = 35;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...ink); doc.text('Filtros aplicados', M, y); y += 5;
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...ink2);
    for (const t of filtersSummary()) { const lines = doc.splitTextToSize(pdfText('• ' + t), W - 2 * M); doc.text(lines, M, y); y += lines.length * 4.2; }
    y += 2;

    // KPIs en cuadrícula
    const kp = state.kpis; const cols = 4; const gw = (W - 2 * M - (cols - 1) * 4) / cols; const gh = 20;
    kp.forEach((k, i) => {
      const cx = M + (i % cols) * (gw + 4), cy = y + Math.floor(i / cols) * (gh + 4);
      doc.setDrawColor(225, 224, 217); doc.setFillColor(249, 249, 247); doc.roundedRect(cx, cy, gw, gh, 2, 2, 'FD');
      doc.setFontSize(8); doc.setTextColor(...ink2); doc.text(pdfText(k.label), cx + 3, cy + 5);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.setTextColor(...ink); doc.text(pdfText(k.value), cx + 3, cy + 12);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(137, 135, 129); doc.text(doc.splitTextToSize(pdfText(k.note), gw - 6)[0], cx + 3, cy + 17);
    });
    y += Math.ceil(kp.length / cols) * (gh + 4) + 2;

    // Hallazgos plazo–monto
    const c = state.corr || {};
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...ink); doc.text(pdfText('Relación plazo – monto'), M, y); y += 5;
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...ink2);
    doc.text(pdfText(`Pearson r = ${Number.isFinite(c.r) ? nf2.format(c.r) : '—'} · R² = ${Number.isFinite(c.r2) ? pct(c.r2) : '—'} · Spearman ρ = ${Number.isFinite(c.rho) ? nf2.format(c.rho) : '—'} · Pendiente = ${c.fit ? nf1.format(c.fit.b) + ' mil M por año de plazo' : '—'} · n = ${c.n ?? 0}`), M, y); y += 5;
    for (const t of state.insight) { const lines = doc.splitTextToSize(pdfText(t), W - 2 * M); ensure(lines.length * 4.2); doc.text(lines, M, y); y += lines.length * 4.2 + 1; }

    // Cómo leer este reporte: explicaciones sencillas de indicadores y gráficas, a dos columnas
    {
      newPage();
      const m = metrics();
      doc.setFont('helvetica', 'bold'); doc.setFontSize(15); doc.setTextColor(...ink);
      doc.text(pdfText('Cómo leer este reporte'), M, y + 4); y += 10;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...ink2);
      const intro = doc.splitTextToSize(pdfText('Explicación en palabras sencillas de cada indicador y gráfica. Los ejemplos imaginan que todo lo que el Departamento ha pedido prestado son 10 manzanas, y se calcularon con los datos y filtros de este reporte.'), W - 2 * M);
      doc.text(intro, M, y); y += intro.length * 4 + 3;
      const gw = (W - 2 * M - 8) / 2, LH = 3.5, FS = 8;
      let top = y; // borde superior de las columnas en la página actual
      let gc = 0;
      const gx = () => M + gc * (gw + 8);
      const place = (hgt) => {
        if (y + hgt <= H - M - 4) return;
        if (gc === 0) { gc = 1; y = top; } else { newPage(); gc = 0; y = 16; top = 16; }
      };
      // Párrafo con etiqueta en negrilla al inicio
      const para = (label, text, width) => {
        doc.setFont('helvetica', 'normal'); doc.setFontSize(FS);
        const lines = doc.splitTextToSize(pdfText(`${label} ${text}`), width);
        return { lines, label: pdfText(label), h: lines.length * LH };
      };
      const drawPara = (p, x, y0) => {
        p.lines.forEach((ln, i) => {
          if (i === 0 && ln.startsWith(p.label)) {
            doc.setFont('helvetica', 'bold'); doc.text(p.label, x, y0);
            const lw = doc.getTextWidth(p.label + ' ');
            doc.setFont('helvetica', 'normal'); doc.text(ln.slice(p.label.length).trimStart(), x + lw, y0);
          } else { doc.setFont('helvetica', 'normal'); doc.text(ln, x, y0 + i * LH); }
        });
      };
      const groups = [
        ['Indicadores principales', ['total', 'wterm', 'wrem', 'balance', 'avg', 'mods', 'biggest']],
        ['Relación plazo – monto', ['pearson', 'r2', 'spearman', 'slope', 'annualTotal']],
        ['Gráficas', ['scatter', 'buckets', 'ratio', 'lenderTerm', 'lenders', 'years', 'gantt', 'maturity', 'amort', 'balanceChart', 'modsChart']],
      ];
      for (const [gname, keys] of groups) {
        place(16);
        doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...blue);
        doc.text(pdfText(gname), gx(), y + 3); y += 7;
        for (const key of keys) {
          const d = INFO[key](m);
          doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5);
          const tl = doc.splitTextToSize(pdfText(d.t), gw);
          // En papel no aplican las instrucciones de interacción del tablero
          const paper = (t) => t.split(/(?<=\.)\s+/).filter((x) => !/^(Toque|Tóquela|Con «Log»|La barra de colores de abajo)/.test(x)).join(' ');
          const pw = para('Qué muestra:', d.what, gw), ph = para('Cómo leerlo:', paper(d.how), gw), pe = d.ex ? para('Ejemplo con manzanas:', d.ex, gw - 6) : null;
          const hgt = tl.length * 4.2 + 1 + pw.h + 1.5 + ph.h + (pe ? 3 + pe.h + 2 : 0) + 4;
          place(hgt);
          const x = gx();
          doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5); doc.setTextColor(...ink);
          doc.text(tl, x, y + 3); y += tl.length * 4.2 + 1;
          doc.setTextColor(...ink2); doc.setFontSize(FS);
          drawPara(pw, x, y + 3); y += pw.h + 1.5;
          drawPara(ph, x, y + 3); y += ph.h;
          if (pe) {
            y += 1.5;
            doc.setFillColor(234, 241, 251); doc.rect(x, y, gw, pe.h + 2.5, 'F');
            doc.setFillColor(...blue); doc.rect(x, y, 0.8, pe.h + 2.5, 'F');
            doc.setTextColor(...ink); drawPara(pe, x + 3.5, y + 3.3); y += pe.h + 2.5;
          }
          y += 4;
        }
        y += 2;
      }
    }

    // Gráficas: dos por fila
    const order = [['scatter', 'Plazo (años) vs. monto contratado'], ['buckets', 'Monto por rango de plazo'], ['ratio', 'Carga anual implícita por contrato (mil M/año)'], ['lenderTerm', 'Plazo por acreedor (rango y ponderado)'],
      ['lenders', 'Monto contratado por acreedor'], ['years', 'Monto firmado por año'], ['gantt', 'Cronograma de vigencia'], ['maturity', 'Monto por año de vencimiento'], ['amort', 'Amortización anual estimada (lineal)'], ['balance', 'Saldo teórico de capital'], ['mods', 'Contratos según número de modificaciones']];
    newPage();
    const colW = (W - 2 * M - 8) / 2; let col = 0; let rowH = 0;
    for (const [id, title] of order) {
      const im = imgs[id]; if (!im) continue;
      let w = colW, hgt = w * (im.h / im.w);
      const maxH = H - 2 * M - 14; if (hgt > maxH) { hgt = maxH; w = hgt * (im.w / im.h); }
      if (col === 0) ensure(hgt + 15); else if (y + hgt + 15 > H - M) { newPage(); col = 0; rowH = 0; }
      const x = M + col * (colW + 8);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...ink); doc.text(pdfText(title), x, y + 4);
      let lx = x; const ly = y + 8.5; let off = 6;
      if (im.legend) {
        doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(...ink2);
        for (const [name, col] of im.legend) { doc.setFillColor(...hexRgb(col)); doc.rect(lx, ly - 2.2, 2.4, 2.4, 'F'); doc.text(pdfText(name), lx + 3.5, ly); lx += 3.5 + doc.getTextWidth(pdfText(name)) + 5; }
        off = 11;
      }
      doc.addImage(im.src, 'PNG', x, y + off, w, hgt, undefined, 'FAST');
      rowH = Math.max(rowH, hgt + off + 4);
      if (col === 1) { y += rowH + 2; col = 0; rowH = 0; } else col = 1;
    }
    if (col === 1) y += rowH + 2;

    // Tabla de contratos
    newPage();
    doc.setFont('helvetica', 'bold'); doc.setFontSize(12); doc.setTextColor(...ink); doc.text('Relación de contratos', M, y + 2); y += 6;
    const tRows = sortedTableRows(state.rows);
    doc.autoTable({
      startY: y, margin: { left: M, right: M, top: 14 }, theme: 'grid',
      styles: { font: 'helvetica', fontSize: 7, cellPadding: 1.5, lineColor: [225, 224, 217], lineWidth: 0.1, textColor: ink, overflow: 'linebreak' },
      headStyles: { fillColor: [243, 242, 238], textColor: ink2, fontStyle: 'bold' },
      columnStyles: { 0: { cellWidth: 26 }, 1: { cellWidth: 24 }, 2: { halign: 'right', cellWidth: 26 }, 3: { cellWidth: 'auto' }, 4: { cellWidth: 16 }, 5: { cellWidth: 16 }, 6: { cellWidth: 16 }, 7: { halign: 'right', cellWidth: 11 }, 8: { cellWidth: 17 }, 9: { cellWidth: 14, textColor: blue } },
      head: [['N.º contrato', 'Acreedor', 'Valor (COP)', 'Objeto', 'Firma', 'Inicio', 'Fin', 'Plazo', 'Estado', 'SECOP']],
      body: tRows.map((r) => [r.ref, r.lender, fmtCOP(r.value), r.object.length > 220 ? r.object.slice(0, 217) + '…' : r.object, fmtDate(r.signed), fmtDate(r.start), fmtDate(r.end), Number.isFinite(r.term) ? nf1.format(r.term) : '—', r.status, r.url ? 'Abrir' : '—'].map(pdfText)),
      foot: [[`Total (${tRows.length})`, '', fmtCOP(sum(tRows, (r) => r.value)), '', '', '', '', nf1.format(wavg(tRows, (r) => r.term) || 0), '', ''].map(pdfText)],
      footStyles: { fillColor: [243, 242, 238], textColor: ink, fontStyle: 'bold' },
      didDrawCell: (d) => { if (d.section === 'body' && d.column.index === 9 && tRows[d.row.index]?.url) doc.link(d.cell.x, d.cell.y, d.cell.width, d.cell.height, { url: tRows[d.row.index].url }); },
      didDrawPage: () => header(),
    });

    // Tablas resumen
    for (const id of ['lenders', 'buckets', 'years', 'maturity']) {
      const s = state.specs[id]; if (!s) continue;
      y = doc.lastAutoTable.finalY + 8; ensure(30);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...ink); doc.text(pdfText(s.title), M, y); y += 3;
      doc.autoTable({
        startY: y, margin: { left: M, right: M, top: 14 }, theme: 'grid', tableWidth: 'wrap',
        styles: { font: 'helvetica', fontSize: 8, cellPadding: 1.5, lineColor: [225, 224, 217], lineWidth: 0.1, textColor: ink },
        headStyles: { fillColor: [243, 242, 238], textColor: ink2, fontStyle: 'bold' },
        columnStyles: Object.fromEntries(s.columns.map((c2, i) => [i, { halign: c2.num ? 'right' : 'left' }])),
        head: [s.columns.map((c2) => pdfText(c2.label))],
        body: s.rows.map((r) => r.map((v, i) => pdfText(s.columns[i].fmt ? s.columns[i].fmt(v) : v))),
        didDrawPage: () => header(),
      });
    }

    // Notas metodológicas y numeración
    y = doc.lastAutoTable.finalY + 8; ensure(40);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...ink); doc.text(pdfText('Notas metodológicas'), M, y); y += 5;
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...ink2);
    for (const li of document.querySelectorAll('#methodology li')) { const lines = doc.splitTextToSize(pdfText('• ' + li.textContent), W - 2 * M); ensure(lines.length * 3.8); doc.text(lines, M, y); y += lines.length * 3.8 + 0.8; }
    const pages = doc.getNumberOfPages();
    for (let i = 1; i <= pages; i++) { doc.setPage(i); doc.setFontSize(8); doc.setTextColor(137, 135, 129); doc.text(pdfText(`Página ${i} de ${pages}`), W - M, H - 6, { align: 'right' }); }
    doc.save(`reporte_credito_publico_valle_${stamp()}.pdf`);
    $('#busy').hidden = true;
  }

  /* ------------------------------------------------------------------- init */
  function init() {
    const saved = store.get(LS_THEME);
    if (saved === 'light' || saved === 'dark') document.documentElement.setAttribute('data-theme', saved);
    readTokens();
    loadColumns();
    renderColMenu();
    syncSortMobile();
    wireControls();
    addChartInfoButtons();
    document.addEventListener('click', (e) => { const b = e.target.closest && e.target.closest('.info-btn'); if (b && state.all.length) { e.preventDefault(); e.stopPropagation(); openInfo(b.dataset.info); } }, true);
    fetch(DATA_URL, { cache: 'no-cache' })
      .then((r) => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.arrayBuffer(); })
      .then((buf) => loadWorkbook(buf, DATA_URL.split('/').pop()))
      .catch((err) => showNotice(`No se pudo cargar ${DATA_URL} (${err.message}). Si abrió el archivo localmente, publíquelo en un servidor o cargue el Excel manualmente.`, true));
  }
  init();
})();
