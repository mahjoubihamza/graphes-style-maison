/* Moteur de mise en page des graphes « style maison ».
 * Chaque graphe est décrit par une grille de données (tableau de chaînes, comme une feuille Excel)
 * et des options. La mise en page produit une liste de primitives en points (rect, text, line,
 * arrow, bubble) dans une boîte W x H. Les mêmes primitives servent à l'aperçu SVG du volet
 * et au dessin en formes PowerPoint (render-office.js).
 */
(function (root) {
  'use strict';

  const PAL = {
    accent: '#1F4E79',
    greys: ['#7F8A96', '#DDE1E5', '#B3BAC2', '#5B636C', '#EEF0F2', '#9AA3AD'],
    ink: '#1A1A1A', mute: '#5B636C', faint: '#7A828B', line: '#333333', white: '#FFFFFF',
  };
  const NB = ' ';
  const MINUS = '−';

  /* ---------- Nombres au format français ---------- */
  function fmt(v, d) {
    d = d || 0;
    const s = Math.abs(v).toFixed(d).split('.');
    const i = s[0].replace(/\B(?=(\d{3})+(?!\d))/g, NB);
    const r = i + (s[1] ? ',' + s[1] : '');
    return (v < 0 && Number(r.replace(/\D/g, '')) !== 0 ? MINUS : '') + r;
  }
  function signed(v, d) {
    const r = fmt(Math.abs(v), d);
    if (Number(r.replace(/\D/g, '')) === 0) return '±' + r;
    return (v > 0 ? '+' : MINUS) + r;
  }
  function parseNum(s) {
    if (s === null || s === undefined) return null;
    if (typeof s === 'number') return isFinite(s) ? s : null;
    let t = String(s).trim().replace(/[\s  %€$]/g, '').replace(MINUS, '-');
    if (t === '') return null;
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) t = t.replace(/\./g, ''); // 1.234,5
    t = t.replace(',', '.');
    const n = Number(t);
    return isFinite(n) ? n : null;
  }
  const withUnit = (s, unit) => (unit ? s + NB + unit : s);

  /* ---------- Outils de mise en page ---------- */
  const textW = (s, fs, bold) => String(s).length * fs * (bold ? 0.6 : 0.55);
  function lum(hex) {
    const n = parseInt(hex.replace('#', ''), 16);
    const c = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map(v => {
      v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
    });
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
  }
  const onColor = hex => (lum(hex) < 0.35 ? PAL.white : PAL.ink);

  function T(cx, cy, text, o) {
    o = o || {};
    const fs = o.size, lines = String(text).split('\n');
    const w = Math.max(...lines.map(l => textW(l, fs, o.bold))) + fs * 0.8;
    const h = fs * 1.25 * lines.length + fs * 0.2;
    const left = o.align === 'left' ? cx : o.align === 'right' ? cx - w : cx - w / 2;
    return { type: 'text', x: left, y: cy - h / 2, w, h, text: String(text), align: o.align || 'center',
      size: fs, color: o.color || PAL.ink, bold: !!o.bold };
  }
  const R = (x, y, w, h, fill) => ({ type: 'rect', x, y, w, h, fill });
  const Ln = (x1, y1, x2, y2, o) => ({ type: 'line', x1, y1, x2, y2, color: (o && o.color) || PAL.line, weight: (o && o.weight) || 0.75 });
  function bubble(cx, cy, text, fs) {
    const w = textW(text, fs, true) + fs * 2.2, h = fs * 2;
    return { type: 'bubble', x: cx - w / 2, y: cy - h / 2, w, h, text, size: fs, color: PAL.ink, fill: PAL.white, stroke: PAL.line };
  }
  /* Flèche en équerre : monte depuis (xa, ya), passe à yL, redescend sur (xb, yb). */
  function elbowArrow(xa, ya, xb, yb, yL, fs) {
    const s = fs * 0.55;
    return [Ln(xa, ya, xa, yL), Ln(xa, yL, xb, yL), Ln(xb, yL, xb, yb - s),
      { type: 'arrow', x: xb, y: yb, size: s, color: PAL.line }];
  }
  /* Écarte verticalement des étiquettes qui se chevauchent (en gardant l'ordre). */
  function spread(labels, minGap, lo, hi) {
    labels.sort((a, b) => a.cy - b.cy);
    for (let i = 1; i < labels.length; i++)
      if (labels[i].cy - labels[i - 1].cy < minGap) labels[i].cy = labels[i - 1].cy + minGap;
    if (labels.length && labels[labels.length - 1].cy > hi) {
      labels[labels.length - 1].cy = hi;
      for (let i = labels.length - 2; i >= 0; i--)
        if (labels[i + 1].cy - labels[i].cy < minGap) labels[i].cy = labels[i + 1].cy - minGap;
    }
    if (labels.length && labels[0].cy < lo) labels.forEach(l => { l.cy += lo - labels[0].cy; });
    return labels;
  }
  function wrap2(s, maxChars) {
    s = String(s);
    if (s.length <= maxChars || s.indexOf(' ') < 0) return s;
    const p = s.split(' ');
    let best = 1, bestLen = Infinity;
    for (let k = 1; k < p.length; k++) {
      const m = Math.max(p.slice(0, k).join(' ').length, p.slice(k).join(' ').length);
      if (m < bestLen) { bestLen = m; best = k; }
    }
    return p.slice(0, best).join(' ') + '\n' + p.slice(best).join(' ');
  }
  function colorsFor(n, accentIdx, accent) {
    let g = 0;
    return Array.from({ length: n }, (_, i) => (i === accentIdx ? accent : PAL.greys[(g++) % PAL.greys.length]));
  }

  /* ================================================================
   * 1. Barres empilées avec totaux et bulle de TCAM
   * Grille : ligne 0 = ['', cat1, cat2, …] ; lignes suivantes = [série, v1, v2, …]
   * ================================================================ */
  const stacked = {
    id: 'stacked',
    label: 'Barres empilées',
    defaultGrid: [
      ['', '2020', '2021', '2022', '2023', '2024', '2025'],
      ['Produits', '120', '124', '128', '131', '133', '136'],
      ['Licences', '30', '32', '35', '37', '40', '42'],
      ['Services', '60', '72', '85', '98', '112', '128'],
    ],
    defaultOptions: { unit: '', decimals: 0, accent: 'Services', showTotals: true, showCagr: true,
      cagrFrom: 0, cagrTo: -1, seriesLabels: true, barWidth: 0.72 },
    gridHint: 'Ligne 1 : catégories (années). Colonne A : séries. Collez directement depuis Excel.',
    parse(grid) {
      const cats = grid[0].slice(1).map(String);
      const series = grid.slice(1).filter(r => String(r[0] || '').trim() !== '' || r.slice(1).some(c => parseNum(c) !== null))
        .map(r => ({ name: String(r[0] || '').trim(), values: cats.map((_, i) => Math.max(0, parseNum(r[i + 1]) || 0)) }));
      // Retire les catégories entièrement vides en fin de grille
      let n = cats.length;
      while (n > 0 && cats[n - 1].trim() === '' && series.every(s => !s.values[n - 1])) n--;
      return { cats: cats.slice(0, n), series: series.map(s => ({ name: s.name, values: s.values.slice(0, n) })) };
    },
    accentChoices(grid) { return this.parse(grid).series.map(s => s.name); },
    layout(m, o, W, H) {
      const fs = o.fontSize, P = [], n = m.cats.length;
      if (!n || !m.series.length) return P;
      const accentIdx = m.series.findIndex(s => s.name === o.accent);
      const cols = colorsFor(m.series.length, accentIdx, o.accentColor);
      const labelW = o.seriesLabels ? Math.min(W * 0.25, Math.max(...m.series.map((s, i) => textW(s.name, fs, i === accentIdx))) + fs * 1.2) : 0;
      const plotW = W - labelW, slot = plotW / n, bw = slot * o.barWidth, cx = i => slot * (i + 0.5);
      const base = H - fs * 2;
      const topRes = o.showCagr && n > 1 ? fs * 4.2 : fs * 0.3;
      const totRes = o.showTotals ? fs * 1.6 : fs * 0.4;
      const tot = m.cats.map((_, i) => m.series.reduce((a, s) => a + s.values[i], 0));
      const mx = Math.max(...tot) || 1, sc = (base - topRes - totRes) / mx;
      const topOf = i => base - tot[i] * sc;
      const sideLabels = [];
      m.cats.forEach((cat, i) => {
        let y = base;
        m.series.forEach((s, k) => {
          const h = s.values[i] * sc;
          if (h > 0) {
            P.push(R(cx(i) - bw / 2, y - h, bw, h, cols[k]));
            const lab = fmt(s.values[i], o.decimals);
            if (h >= fs * 1.3 && textW(lab, fs) <= bw - 4)
              P.push(T(cx(i), y - h / 2, lab, { size: fs, color: onColor(cols[k]), bold: k === accentIdx }));
          }
          if (i === n - 1 && o.seriesLabels)
            sideLabels.push({ cy: y - h / 2, name: s.name, k });
          y -= h;
        });
        if (o.showTotals) P.push(T(cx(i), topOf(i) - fs * 0.85, fmt(tot[i], o.decimals), { size: fs, bold: true }));
        P.push(T(cx(i), base + fs * 1.05, cat, { size: fs }));
      });
      spread(sideLabels, fs * 1.3, topRes, base - fs * 0.6).forEach(l =>
        P.push(T(cx(n - 1) + bw / 2 + fs * 0.6, l.cy, l.name, { size: fs, align: 'left',
          color: l.k === accentIdx ? o.accentColor : PAL.mute, bold: l.k === accentIdx })));
      P.push(Ln(0, base, plotW, base, { weight: 0.75 }));
      if (o.showCagr && n > 1) {
        const a = Math.max(0, Math.min(n - 1, o.cagrFrom)), b = o.cagrTo < 0 ? n - 1 : Math.min(n - 1, o.cagrTo);
        if (b > a && tot[a] > 0) {
          const g = (Math.pow(tot[b] / tot[a], 1 / (b - a)) - 1) * 100;
          const off = o.showTotals ? fs * 1.9 : fs * 0.5, yL = fs * 1.1;
          P.push(...elbowArrow(cx(a), topOf(a) - off, cx(b), topOf(b) - off, yL, fs));
          P.push(bubble((cx(a) + cx(b)) / 2, yL, 'TCAM' + NB + signed(g, 1) + NB + '%', fs));
        }
      }
      return P;
    },
  };

  /* ================================================================
   * 2. Waterfall (pont) avec connecteurs et flèche d'écart
   * Grille : ['Libellé', 'Valeur', 'Total (x)'] puis une ligne par marche
   * ================================================================ */
  const waterfall = {
    id: 'waterfall',
    label: 'Waterfall',
    defaultGrid: [
      ['Libellé', 'Valeur', 'Total (x)'],
      ['EBITDA 2024', '48', 'x'],
      ['Volume', '9', ''],
      ['Prix', '6', ''],
      ['Mix produit', '-3', ''],
      ['Coûts matières', '-7', ''],
      ['Coûts de personnel', '-4', ''],
      ['Productivité', '5', ''],
      ['EBITDA 2025', '', 'x'],
    ],
    defaultOptions: { unit: 'M€', decimals: 0, showDiff: true, barWidth: 0.72 },
    gridHint: 'Marquez les totaux d\'un « x ». Un total sans valeur est calculé.',
    parse(grid) {
      return { items: grid.slice(1).filter(r => String(r[0] || '').trim() !== '' || parseNum(r[1]) !== null)
        .map(r => ({ name: String(r[0] || '').trim(), value: parseNum(r[1]), total: /\S/.test(String(r[2] || '')) })) };
    },
    layout(m, o, W, H) {
      const fs = o.fontSize, P = [], n = m.items.length;
      if (!n) return P;
      let lv = 0;
      const bars = m.items.map(it => {
        let b, t, v;
        if (it.total) { v = it.value === null ? lv : it.value; b = Math.min(0, v); t = Math.max(0, v); lv = v; }
        else { v = it.value || 0; b = Math.min(lv, lv + v); t = Math.max(lv, lv + v); lv += v; }
        return { name: it.name, v, total: it.total, b, t, end: lv };
      });
      const mx = Math.max(0, ...bars.map(d => d.t)), mn = Math.min(0, ...bars.map(d => d.b));
      const totals = bars.map((d, i) => (d.total ? i : -1)).filter(i => i >= 0);
      const diff = o.showDiff && totals.length >= 2;
      const maxChars = Math.max(6, Math.floor((W / n) / (fs * 0.55)));
      const twoLines = m.items.some(it => wrap2(it.name, maxChars).indexOf('\n') >= 0);
      const catH = fs * (twoLines ? 3.2 : 2);
      const topRes = diff ? fs * 4.2 : fs * 0.3, valRes = fs * 1.6;
      const base = H - catH - (mn < 0 ? valRes : 0);
      const sc = (base - topRes - valRes) / ((mx - mn) || 1);
      const Y = v => base - (v - mn) * sc;
      const slot = W / n, bw = slot * o.barWidth, cx = i => slot * (i + 0.5);
      bars.forEach((d, i) => {
        const y = Y(d.t), h = Math.max((d.t - d.b) * sc, 0.75);
        const col = d.total ? o.accentColor : d.v >= 0 ? PAL.greys[0] : PAL.greys[2];
        P.push(R(cx(i) - bw / 2, y, bw, h, col));
        const lab = d.total ? fmt(d.v, o.decimals) : signed(d.v, o.decimals);
        if (d.v >= 0) P.push(T(cx(i), y - fs * 0.85, lab, { size: fs, bold: d.total }));
        else P.push(T(cx(i), y + h + fs * 0.85, lab, { size: fs, bold: d.total }));
        if (i < n - 1) P.push(Ln(cx(i) + bw / 2, Y(d.end), cx(i + 1) - bw / 2, Y(d.end), { weight: 0.5, color: '#555555' }));
        P.push(T(cx(i), H - catH + fs * (twoLines ? 1.6 : 1.05), wrap2(d.name, maxChars), { size: fs }));
      });
      P.push(Ln(0, Y(0), W, Y(0), { weight: 0.75 }));
      if (diff) {
        const a = totals[0], b = totals[totals.length - 1], A = bars[a], B = bars[b];
        const yL = fs * 1.1, off = fs * 1.9;
        P.push(...elbowArrow(cx(a), Y(A.t) - off, cx(b), Y(B.t) - off, yL, fs));
        const dv = B.v - A.v;
        let s = withUnit(signed(dv, o.decimals), o.unit);
        if (A.v) s += ' (' + signed(dv / Math.abs(A.v) * 100, 1) + NB + '%)';
        P.push(bubble((cx(a) + cx(b)) / 2, yL, s, fs));
      }
      return P;
    },
  };

  /* ================================================================
   * 3. Mekko (largeur = taille du segment, hauteur = parts en %)
   * Grille : ['Segment', 'Taille', acteur1, acteur2, …] puis une ligne par segment
   * ================================================================ */
  const mekko = {
    id: 'mekko',
    label: 'Mekko',
    defaultGrid: [
      ['Segment', 'Taille', 'Notre groupe', 'Concurrent A', 'Concurrent B', 'Autres'],
      ['Grande distribution', '420', '18', '32', '22', '28'],
      ['Spécialistes', '260', '26', '20', '14', '40'],
      ['E-commerce', '180', '34', '18', '10', '38'],
      ['B2B / Pro', '140', '12', '28', '30', '30'],
    ],
    defaultOptions: { unit: 'M€', decimals: 0, accent: 'Notre groupe', seriesLabels: true, showFooter: true },
    gridHint: 'Colonne B : taille du segment. Colonnes suivantes : parts par acteur (ramenées à 100 %).',
    parse(grid) {
      const players = grid[0].slice(2).map(s => String(s).trim());
      const segs = grid.slice(1).filter(r => String(r[0] || '').trim() !== '' || parseNum(r[1]) !== null).map(r => {
        const raw = players.map((_, i) => Math.max(0, parseNum(r[i + 2]) || 0));
        const sum = raw.reduce((a, b) => a + b, 0) || 1;
        return { name: String(r[0] || '').trim(), size: Math.max(0, parseNum(r[1]) || 0), shares: raw.map(v => v / sum * 100) };
      });
      let np = players.length;
      while (np > 0 && players[np - 1] === '' && segs.every(s => !s.shares[np - 1])) np--;
      return { players: players.slice(0, np), segs: segs.map(s => ({ ...s, shares: s.shares.slice(0, np) })) };
    },
    accentChoices(grid) { return this.parse(grid).players; },
    layout(m, o, W, H) {
      const fs = o.fontSize, P = [];
      if (!m.segs.length || !m.players.length) return P;
      const accentIdx = m.players.indexOf(o.accent);
      const cols = colorsFor(m.players.length, accentIdx, o.accentColor);
      const labelW = o.seriesLabels ? Math.min(W * 0.25, Math.max(...m.players.map((s, i) => textW(s, fs, i === accentIdx))) + fs * 1.2) : 0;
      const tot = m.segs.reduce((a, s) => a + s.size, 0) || 1;
      const gap = Math.max(1.5, fs * 0.2), plotW = W - labelW, av = plotW - gap * (m.segs.length - 1);
      const top = fs * 1.8, footer = o.showFooter ? fs * 1.8 : 0, base = H - fs * 3.4 - footer, Hh = base - top;
      let x = 0;
      const sideLabels = [];
      m.segs.forEach((s, j) => {
        const w = s.size / tot * av;
        let y = base;
        s.shares.forEach((p, i) => {
          const h = p / 100 * Hh;
          if (h > 0) {
            P.push(R(x, y - h, w, h, cols[i]));
            const lab = fmt(p, 0) + NB + '%';
            if (h >= fs * 1.3 && textW(lab, fs) <= w - 4)
              P.push(T(x + w / 2, y - h / 2, lab, { size: fs, color: onColor(cols[i]), bold: i === accentIdx }));
          }
          if (j === m.segs.length - 1 && o.seriesLabels) sideLabels.push({ cy: y - h / 2, name: m.players[i], k: i });
          y -= h;
        });
        const maxChars = Math.max(5, Math.floor(w / (fs * 0.55)));
        P.push(T(x + w / 2, top - fs * 0.85, withUnit(fmt(s.size, o.decimals), o.unit), { size: fs, bold: true }));
        const nm = s.name.length > maxChars ? wrap2(s.name, maxChars) : s.name;
        const nLines = nm.split('\n').length;
        P.push(T(x + w / 2, base + fs * (nLines > 1 ? 1.6 : 1.05), nm, { size: fs }));
        P.push(T(x + w / 2, base + fs * (nLines > 1 ? 3.0 : 2.4), fmt(s.size / tot * 100, 0) + NB + '%', { size: fs * 0.85, color: PAL.faint }));
        x += w + gap;
      });
      spread(sideLabels, fs * 1.3, top + fs * 0.6, base - fs * 0.6).forEach(l =>
        P.push(T(plotW + fs * 0.6, l.cy, l.name, { size: fs, align: 'left',
          color: l.k === accentIdx ? o.accentColor : PAL.mute, bold: l.k === accentIdx })));
      if (o.showFooter) {
        let s = 'Total marché : ' + withUnit(fmt(tot, o.decimals), o.unit);
        if (accentIdx >= 0) s += ' · part ' + (m.players[accentIdx]) + ' : ' + fmt(m.segs.reduce((a, g) => a + g.size * g.shares[accentIdx], 0) / tot, 0) + NB + '%';
        P.push(T(0, H - footer / 2, s, { size: fs * 0.85, align: 'left', color: PAL.mute }));
      }
      return P;
    },
  };

  const CHARTS = { stacked, waterfall, mekko };
  const BASE_OPTIONS = { fontSize: 12, fontName: 'Arial', accentColor: PAL.accent };

  function build(spec) {
    const c = CHARTS[spec.type];
    const o = Object.assign({}, BASE_OPTIONS, c.defaultOptions, spec.options || {});
    return c.layout(c.parse(spec.grid), o, spec.box.width, spec.box.height);
  }
  function newSpec(type) {
    const c = CHARTS[type];
    return { v: 1, type, grid: c.defaultGrid.map(r => r.slice()), options: Object.assign({}, BASE_OPTIONS, c.defaultOptions),
      box: { left: 40, top: 130, width: 880, height: 340 } };
  }

  /* ---------- Rendu SVG (aperçu du volet et tests) ---------- */
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  function toSVG(P, W, H, fontName) {
    const ff = esc((fontName || 'Arial') + ',Helvetica,sans-serif');
    const pad = 4;
    let o = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${-pad} ${-pad} ${W + 2 * pad} ${H + 2 * pad}" font-family="${ff}">`;
    o += `<rect x="${-pad}" y="${-pad}" width="${W + 2 * pad}" height="${H + 2 * pad}" fill="#fff"/>`;
    const txt = (p, x, cy) => {
      const lines = p.text.split('\n'), lh = p.size * 1.25, y0 = cy - lh * (lines.length - 1) / 2;
      const anchor = p.align === 'left' ? 'start' : p.align === 'right' ? 'end' : 'middle';
      return lines.map((l, i) => `<text x="${x}" y="${y0 + i * lh}" font-size="${p.size}" fill="${p.color}" font-weight="${p.bold ? 700 : 400}" text-anchor="${anchor}" dominant-baseline="central">${esc(l)}</text>`).join('');
    };
    for (const p of P) {
      if (p.type === 'rect') o += `<rect x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" fill="${p.fill}" stroke="#fff" stroke-width="0.75"/>`;
      else if (p.type === 'line') o += `<line x1="${p.x1}" y1="${p.y1}" x2="${p.x2}" y2="${p.y2}" stroke="${p.color}" stroke-width="${p.weight}"/>`;
      else if (p.type === 'arrow') o += `<path d="M${p.x - p.size / 2},${p.y - p.size} L${p.x + p.size / 2},${p.y - p.size} L${p.x},${p.y} z" fill="${p.color}"/>`;
      else if (p.type === 'bubble') o += `<rect x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" rx="${p.h / 2}" fill="${p.fill}" stroke="${p.stroke}" stroke-width="0.75"/>` + txt({ ...p, bold: true, align: 'center' }, p.x + p.w / 2, p.y + p.h / 2);
      else if (p.type === 'text') o += txt(p, p.align === 'left' ? p.x : p.align === 'right' ? p.x + p.w : p.x + p.w / 2, p.y + p.h / 2);
    }
    return o + '</svg>';
  }

  const api = { PAL, CHARTS, BASE_OPTIONS, build, newSpec, toSVG, fmt, signed, parseNum };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.TCL = api;
})(typeof self !== 'undefined' ? self : this);
