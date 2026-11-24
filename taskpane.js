/* Volet de l'add-in : choix du graphe, tableau de données, options, aperçu, insertion et mise à jour. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const inOffice = () => typeof Office !== 'undefined' && Office.context && Office.context.host;
  let spec = null;          // graphe en cours d'édition
  let editing = null;       // { spec, shape, members } si un graphe existant est sélectionné
  const byType = {};        // garde les données saisies par type quand on passe de l'un à l'autre

  /* ---------- Options affichées selon le type ---------- */
  const OPTS = {
    stacked: [
      { k: 'accent', label: 'Série mise en avant', type: 'accent' },
      { k: 'accentColor', label: 'Couleur d\'accent', type: 'color' },
      { k: 'unit', label: 'Unité', type: 'text' },
      { k: 'decimals', label: 'Décimales', type: 'number', min: 0, max: 3 },
      { k: 'cagrFrom', label: 'TCAM depuis', type: 'cat' },
      { k: 'cagrTo', label: 'TCAM jusqu\'à', type: 'cat', last: true },
      { k: 'showCagr', label: 'Bulle de TCAM', type: 'check' },
      { k: 'showTotals', label: 'Totaux', type: 'check' },
      { k: 'seriesLabels', label: 'Noms des séries à droite', type: 'check' },
      { k: 'fontSize', label: 'Taille du texte (pt)', type: 'number', min: 8, max: 20 },
    ],
    waterfall: [
      { k: 'accentColor', label: 'Couleur des totaux', type: 'color' },
      { k: 'unit', label: 'Unité', type: 'text' },
      { k: 'decimals', label: 'Décimales', type: 'number', min: 0, max: 3 },
      { k: 'fontSize', label: 'Taille du texte (pt)', type: 'number', min: 8, max: 20 },
      { k: 'showDiff', label: 'Flèche d\'écart entre totaux', type: 'check' },
    ],
    mekko: [
      { k: 'accent', label: 'Acteur mis en avant', type: 'accent' },
      { k: 'accentColor', label: 'Couleur d\'accent', type: 'color' },
      { k: 'unit', label: 'Unité des tailles', type: 'text' },
      { k: 'decimals', label: 'Décimales', type: 'number', min: 0, max: 3 },
      { k: 'seriesLabels', label: 'Noms des acteurs à droite', type: 'check' },
      { k: 'showFooter', label: 'Ligne de total en pied', type: 'check' },
      { k: 'fontSize', label: 'Taille du texte (pt)', type: 'number', min: 8, max: 20 },
    ],
  };

  function setStatus(msg, err) { $('status').textContent = msg || ''; $('status').className = 'status' + (err ? ' err' : ''); }

  /* ---------- Rendu de l'interface ---------- */
  function renderAll() {
    document.querySelectorAll('.types button').forEach(b => b.setAttribute('aria-selected', b.dataset.type === spec.type));
    $('hint').textContent = TCL.CHARTS[spec.type].gridHint;
    renderGrid();
    renderOptions();
    renderPreview();
    renderMode();
  }
  function renderMode() {
    const ed = !!editing;
    $('mode').innerHTML = ed ? 'Graphe sélectionné : <b>modifiez puis mettez à jour</b>' : 'Nouveau graphe';
    $('insert').hidden = ed; $('update').hidden = !ed; $('asNew').hidden = !ed;
  }
  function renderPreview() {
    try {
      const P = TCL.build(spec);
      $('preview').innerHTML = TCL.toSVG(P, spec.box.width, spec.box.height, spec.options.fontName);
    } catch (e) { $('preview').textContent = 'Aperçu impossible : ' + e.message; }
  }

  function renderGrid() {
    const t = $('grid'); t.innerHTML = '';
    const g = spec.grid, nc = Math.max(...g.map(r => r.length));
    g.forEach((row, r) => {
      const tr = document.createElement('tr');
      for (let c = 0; c < nc; c++) {
        const td = document.createElement('td'), inp = document.createElement('input');
        inp.value = row[c] === undefined ? '' : row[c];
        inp.dataset.r = r; inp.dataset.c = c;
        inp.addEventListener('input', () => { setCell(r, c, inp.value); refreshDerived(); });
        inp.addEventListener('paste', onPaste);
        inp.addEventListener('keydown', onKey);
        td.appendChild(inp); tr.appendChild(td);
      }
      t.appendChild(tr);
    });
  }
  function setCell(r, c, v) {
    while (spec.grid.length <= r) spec.grid.push([]);
    const row = spec.grid[r];
    while (row.length <= c) row.push('');
    row[c] = v;
  }
  function onPaste(e) {
    const txt = (e.clipboardData || window.clipboardData).getData('text');
    if (!/[\t\n]/.test(txt)) return;
    e.preventDefault();
    const r0 = +e.target.dataset.r, c0 = +e.target.dataset.c;
    txt.replace(/\r/g, '').replace(/\n$/, '').split('\n').forEach((line, i) =>
      line.split('\t').forEach((v, j) => setCell(r0 + i, c0 + j, v.trim())));
    const nc = Math.max(...spec.grid.map(r => r.length));
    spec.grid.forEach(r => { while (r.length < nc) r.push(''); });
    renderGrid(); refreshDerived();
  }
  function onKey(e) {
    const r = +e.target.dataset.r, c = +e.target.dataset.c;
    let tr = r, tc = c;
    if (e.key === 'Enter' || e.key === 'ArrowDown') tr++;
    else if (e.key === 'ArrowUp') tr--;
    else return;
    const n = document.querySelector(`#grid input[data-r="${tr}"][data-c="${tc}"]`);
    if (n) { e.preventDefault(); n.focus(); n.select(); }
  }
  function refreshDerived() { renderOptions(); renderPreview(); }

  function renderOptions() {
    const box = $('options'), focused = document.activeElement && document.activeElement.dataset.opt;
    box.innerHTML = '';
    const o = spec.options, c = TCL.CHARTS[spec.type];
    OPTS[spec.type].forEach(def => {
      const lab = document.createElement('label');
      let inp;
      if (def.type === 'check') {
        lab.className = 'check';
        inp = document.createElement('input'); inp.type = 'checkbox'; inp.checked = !!o[def.k];
        inp.addEventListener('change', () => { o[def.k] = inp.checked; renderPreview(); });
        lab.append(inp, document.createTextNode(def.label));
      } else {
        lab.append(document.createTextNode(def.label));
        if (def.type === 'accent' || def.type === 'cat') {
          inp = document.createElement('select');
          const items = def.type === 'accent' ? [['', 'Aucun'], ...c.accentChoices(spec.grid).map(s => [s, s])]
            : c.parse(spec.grid).cats.map((s, i) => [String(i), s]);
          items.forEach(([v, t]) => { const op = document.createElement('option'); op.value = v; op.textContent = t || '(sans nom)'; inp.appendChild(op); });
          let cur = o[def.k];
          if (def.type === 'cat') cur = String(cur < 0 ? items.length - 1 : cur);
          inp.value = cur;
          inp.addEventListener('change', () => { o[def.k] = def.type === 'cat' ? +inp.value : inp.value; renderPreview(); });
        } else {
          inp = document.createElement('input'); inp.type = def.type;
          if (def.min !== undefined) { inp.min = def.min; inp.max = def.max; }
          inp.value = o[def.k];
          inp.addEventListener('input', () => {
            o[def.k] = def.type === 'number' ? (Number(inp.value) || 0) : inp.value; renderPreview();
          });
        }
        lab.appendChild(inp);
      }
      inp.dataset.opt = def.k;
      box.appendChild(lab);
    });
    if (focused) { const f = box.querySelector(`[data-opt="${focused}"]`); if (f && f.tagName !== 'SELECT') f.focus(); }
  }

  /* ---------- Actions de la grille ---------- */
  function addRow() { const nc = Math.max(...spec.grid.map(r => r.length)); spec.grid.push(Array(nc).fill('')); renderGrid(); }
  function delRow() { if (spec.grid.length > 2) { spec.grid.pop(); renderGrid(); refreshDerived(); } }
  function addCol() { spec.grid.forEach(r => r.push('')); renderGrid(); }
  function delCol() { if (spec.grid[0].length > 2) { spec.grid.forEach(r => r.pop()); renderGrid(); refreshDerived(); } }

  function switchType(type) {
    if (spec.type === type) return;
    byType[spec.type] = spec;
    const box = spec.box;
    spec = byType[type] || TCL.newSpec(type);
    spec.box = box;
    if (editing) { editing = null; }
    renderAll();
  }

  /* ---------- PowerPoint ---------- */
  async function insert(asNew) {
    if (!inOffice()) { setStatus('Ouvrez ce volet dans PowerPoint pour insérer le graphe.', true); return; }
    setStatus('Dessin en cours…');
    $('insert').disabled = $('update').disabled = true;
    try {
      await PowerPoint.run(async ctx => {
        const slides = ctx.presentation.getSelectedSlides();
        slides.load('items');
        await ctx.sync();
        if (!slides.items.length) throw new Error('Sélectionnez d\'abord une diapositive.');
        const slide = slides.items[0];
        const out = JSON.parse(JSON.stringify(spec));
        delete out._drawn;
        if (editing && !asNew) {
          // Supprime l'ancien graphe et redessine à la même place
          const cur = await TCLOffice.findSelectedChart(ctx);
          if (cur) {
            out.box = TCLOffice.adjustBox(cur.spec, cur.shape);
            (cur.members || [cur.shape]).forEach(s => s.delete());
            await ctx.sync();
          }
        } else if (asNew && editing) {
          out.box = { ...out.box, left: out.box.left + 20, top: out.box.top + 20 };
        }
        const P = TCL.build(out);
        const shape = await TCLOffice.draw(ctx, slide, P, out, out.options.fontName);
        spec = out;
        if (TCLOffice.supports('1.5')) { ctx.presentation.setSelectedShapes([shape.id]); await ctx.sync().catch(() => null); }
      });
      editing = { spec };
      renderMode();
      setStatus('Graphe dessiné. Sélectionnez-le plus tard pour le modifier.');
    } catch (e) {
      console.error(e);
      setStatus('Erreur : ' + (e.message || e), true);
    } finally {
      $('insert').disabled = $('update').disabled = false;
    }
  }

  let selBusy = false;
  async function onSelectionChanged() {
    if (selBusy || !inOffice()) return;
    selBusy = true;
    try {
      await PowerPoint.run(async ctx => {
        const found = await TCLOffice.findSelectedChart(ctx);
        if (found) {
          if (!editing || JSON.stringify(found.spec.grid) !== JSON.stringify(spec.grid) || found.spec.type !== spec.type) {
            spec = found.spec;
            spec.options = Object.assign({}, TCL.BASE_OPTIONS, TCL.CHARTS[spec.type].defaultOptions, spec.options);
          }
          editing = { spec };
          renderAll();
        } else if (editing) {
          editing = null; renderMode();
        }
      });
    } catch (e) { console.warn(e); } finally { selBusy = false; }
  }

  /* ---------- Démarrage ---------- */
  function start(initialType) {
    spec = TCL.newSpec(initialType || 'stacked');
    document.querySelectorAll('.types button').forEach(b => b.addEventListener('click', () => switchType(b.dataset.type)));
    $('addRow').onclick = addRow; $('delRow').onclick = delRow; $('addCol').onclick = addCol; $('delCol').onclick = delCol;
    $('reset').onclick = () => { const box = spec.box; spec = TCL.newSpec(spec.type); spec.box = box; renderAll(); };
    $('insert').onclick = () => insert(false);
    $('update').onclick = () => insert(false);
    $('asNew').onclick = () => insert(true);
    renderAll();
  }
  const qType = new URLSearchParams(location.search).get('type');
  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(info => {
      start(qType);
      if (info && info.host === Office.HostType.PowerPoint) {
        if (!TCLOffice.supports('1.4')) setStatus('Cette version de PowerPoint est trop ancienne (PowerPointApi 1.4 requis).', true);
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
        onSelectionChanged();
      }
    });
  } else {
    document.addEventListener('DOMContentLoaded', () => start(qType));
  }
})();
