/* Dessin des primitives en formes PowerPoint natives (Office.js, PowerPointApi 1.4+).
 * Le graphe est groupé (PowerPointApi 1.8) et porte ses données dans une balise (tag),
 * ce qui permet de le resélectionner et de le recalculer.
 */
(function (root) {
  'use strict';
  const TAG = 'TCLCHART';
  const TAG_ID = 'TCLID';
  const GROUP_NAME = 'Graphe style maison';

  const E = (path, fallback) => {
    try { return path() || fallback; } catch (e) { return fallback; }
  };
  const supports = v => E(() => Office.context.requirements.isSetSupported('PowerPointApi', v), false);

  function noLine(s) { s.lineFormat.visible = false; }
  function styleText(tr, p, fontName) {
    tr.font.name = fontName;
    tr.font.size = p.size;
    tr.font.color = p.color;
    tr.font.bold = !!p.bold;
    tr.paragraphFormat.horizontalAlignment = p.align === 'left' ? E(() => PowerPoint.ParagraphHorizontalAlignment.left, 'Left')
      : p.align === 'right' ? E(() => PowerPoint.ParagraphHorizontalAlignment.right, 'Right')
      : E(() => PowerPoint.ParagraphHorizontalAlignment.center, 'Center');
  }
  function tightFrame(tf) {
    tf.autoSizeSetting = E(() => PowerPoint.ShapeAutoSize.autoSizeNone, 'AutoSizeNone');
    tf.wordWrap = false;
    tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
    tf.verticalAlignment = E(() => PowerPoint.TextVerticalAlignment.middle, 'Middle');
  }

  /* Dessine les primitives sur la diapositive ; renvoie le groupe (ou la 1re forme si pas de groupe). */
  async function draw(context, slide, prims, spec, fontName) {
    const shapes = slide.shapes, box = spec.box, made = [];
    const X = v => box.left + v, Y = v => box.top + v;
    const GT = E(() => PowerPoint.GeometricShapeType, {});
    const chartId = 'tcl' + Date.now().toString(36);
    for (const p of prims) {
      let s;
      if (p.type === 'rect') {
        s = shapes.addGeometricShape(GT.rectangle || 'Rectangle', { left: X(p.x), top: Y(p.y), width: Math.max(p.w, 0.1), height: Math.max(p.h, 0.1) });
        s.fill.setSolidColor(p.fill);
        s.lineFormat.color = '#FFFFFF';
        s.lineFormat.weight = 0.75;
      } else if (p.type === 'text') {
        s = shapes.addTextBox(p.text, { left: X(p.x), top: Y(p.y), width: p.w, height: p.h });
        tightFrame(s.textFrame);
        styleText(s.textFrame.textRange, p, fontName);
        s.fill.clear();
        noLine(s);
      } else if (p.type === 'line') {
        const l = Math.min(p.x1, p.x2), t = Math.min(p.y1, p.y2);
        s = shapes.addLine(E(() => PowerPoint.ConnectorType.straight, 'Straight'),
          { left: X(l), top: Y(t), width: Math.abs(p.x2 - p.x1), height: Math.abs(p.y2 - p.y1) });
        s.lineFormat.color = p.color;
        s.lineFormat.weight = p.weight;
      } else if (p.type === 'arrow') {
        // Triangle pointe en bas (forme « fusion » des organigrammes), pas besoin de rotation
        s = shapes.addGeometricShape(GT.flowChartMerge || 'FlowChartMerge', { left: X(p.x - p.size / 2), top: Y(p.y - p.size), width: p.size, height: p.size });
        s.fill.setSolidColor(p.color);
        noLine(s);
      } else if (p.type === 'bubble') {
        // Pastille arrondie (forme « terminaison ») avec le texte dedans
        s = shapes.addGeometricShape(GT.flowChartTerminator || 'FlowChartTerminator', { left: X(p.x), top: Y(p.y), width: p.w, height: p.h });
        s.fill.setSolidColor(p.fill);
        s.lineFormat.color = p.stroke;
        s.lineFormat.weight = 0.75;
        s.textFrame.textRange.text = p.text;
        tightFrame(s.textFrame);
        styleText(s.textFrame.textRange, { ...p, bold: true, align: 'center' }, fontName);
      }
      if (s) { s.name = 'TCL ' + p.type; made.push(s); }
    }
    made.forEach(s => s.load('id'));
    await context.sync();

    const json = JSON.stringify(spec);
    if (supports('1.8') && made.length > 1) {
      const g = shapes.addGroup(made.map(s => s.id));
      g.name = GROUP_NAME;
      g.tags.add(TAG, json);
      g.load('id,left,top,width,height');
      await context.sync();
      spec._drawn = { left: g.left, top: g.top, width: g.width, height: g.height };
      g.tags.add(TAG, JSON.stringify(spec));
      await context.sync();
      return g;
    }
    // Repli sans groupes : chaque forme porte l'identifiant du graphe, la première porte les données
    made.forEach(s => s.tags.add(TAG_ID, chartId));
    made[0].tags.add(TAG, json);
    await context.sync();
    return made[0];
  }

  /* Cherche un graphe « style maison » dans la sélection. Renvoie { spec, shape, members } ou null. */
  async function findSelectedChart(context) {
    if (!supports('1.5')) return null;
    const sel = context.presentation.getSelectedShapes();
    sel.load('items/id,items/type');
    await context.sync();
    const readTags = async c => {
      const t = c.tags.getItemOrNullObject(TAG); t.load('value');
      const i = c.tags.getItemOrNullObject(TAG_ID); i.load('value');
      c.load('id,left,top,width,height');
      await context.sync();
      return { c, t: t.isNullObject ? null : t.value, i: i.isNullObject ? null : i.value };
    };
    for (const s of sel.items) {
      const found = [await readTags(s)];
      // Une forme à l'intérieur du groupe : remonter au groupe parent
      if (!found[0].t && supports('1.8') && s.type !== 'Group') {
        try { found.push(await readTags(s.parentGroup)); } catch (e) { /* pas dans un groupe */ }
      }
      for (const { c, t, i } of found) {
        if (t) return { spec: JSON.parse(t), shape: c, chartId: i };
        if (i) {
          // Repli sans groupes : retrouver la forme porteuse des données sur la diapositive
          const slide = context.presentation.getSelectedSlides().getItemAt(0);
          const all = slide.shapes; all.load('items/id');
          await context.sync();
          const pairs = all.items.map(x => {
            const a = x.tags.getItemOrNullObject(TAG_ID); a.load('value');
            const b = x.tags.getItemOrNullObject(TAG); b.load('value');
            x.load('left,top,width,height');
            return { x, a, b };
          });
          await context.sync();
          const members = pairs.filter(p => !p.a.isNullObject && p.a.value === i);
          const owner = members.find(p => !p.b.isNullObject);
          if (owner) return { spec: JSON.parse(owner.b.value), shape: null, members: members.map(p => p.x), chartId: i };
        }
      }
    }
    return null;
  }

  /* Si l'utilisateur a déplacé ou redimensionné le groupe, on en tient compte avant de redessiner. */
  function adjustBox(spec, shape) {
    const d = spec._drawn;
    if (!d || !shape || shape.width === undefined) return spec.box;
    const kx = shape.width / d.width, ky = shape.height / d.height;
    const b = spec.box;
    return {
      left: shape.left + (b.left - d.left) * kx,
      top: shape.top + (b.top - d.top) * ky,
      width: b.width * kx,
      height: b.height * ky,
    };
  }

  root.TCLOffice = { draw, findSelectedChart, adjustBox, supports, TAG };
})(typeof self !== 'undefined' ? self : this);
