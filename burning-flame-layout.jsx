// Burning Flame — presentation layout + neck-sleeve pattern
// Run in Illustrator with BURNING FLAMES.ai open: File > Scripts > Other Script...
// Non-destructive: original art is only duplicated, never moved; result is saved under a NEW name.
#target illustrator
(function () {
  if (app.documents.length === 0) { alert('Open BURNING FLAMES.ai first.'); return; }
  var doc = app.activeDocument;
  app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;

  function rgb(r, g, b) { var c = new RGBColor(); c.red = r; c.green = g; c.blue = b; return c; }
  var CREAM = rgb(253, 245, 224), INK = rgb(31, 18, 14);
  var FLAV = [
    { key: 'original', name: 'ORIGINAL', col: rgb(237, 98, 43), heat: '2 / 3' },
    { key: 'mild', name: 'MILD', col: rgb(239, 123, 144), heat: '1 / 3' },
    { key: 'extrahot', name: 'EXTRA HOT', col: rgb(209, 62, 54), heat: '3 / 3' }
  ];

  // ---------- 1. cluster the existing top-level art by overlapping bounds
  var src = doc.layers[0];
  var items = [];
  for (var i = 0; i < src.pageItems.length; i++) {
    var it = src.pageItems[i];
    if (it.parent !== src || it.locked || it.hidden) continue;
    items.push(it);
  }
  function vb(it) { return it.visibleBounds; } // [L, T, R, B]
  function overlap(a, b, pad) { return !(a[2] + pad < b[0] || b[2] + pad < a[0] || a[3] - pad > b[1] || b[3] - pad > a[1]); }
  var clusters = [];
  for (i = 0; i < items.length; i++) clusters.push({ items: [items[i]], b: vb(items[i]).slice(0) });
  var merged = true;
  while (merged) {
    merged = false;
    outer: for (var a = 0; a < clusters.length; a++) {
      for (var b2 = a + 1; b2 < clusters.length; b2++) {
        if (overlap(clusters[a].b, clusters[b2].b, 0.5)) {
          var A = clusters[a], B = clusters[b2];
          A.items = A.items.concat(B.items);
          A.b = [Math.min(A.b[0], B.b[0]), Math.max(A.b[1], B.b[1]), Math.max(A.b[2], B.b[2]), Math.min(A.b[3], B.b[3])];
          clusters.splice(b2, 1); merged = true; break outer;
        }
      }
    }
  }
  function W(c) { return c.b[2] - c.b[0]; } function H(c) { return c.b[1] - c.b[3]; }
  function cx(c) { return (c.b[0] + c.b[2]) / 2; } function cy(c) { return (c.b[1] + c.b[3]) / 2; }
  var maxH = 0; for (i = 0; i < clusters.length; i++) maxH = Math.max(maxH, H(clusters[i]));
  var big = [], labels = [], logos = [];
  for (i = 0; i < clusters.length; i++) if (H(clusters[i]) > maxH * 0.55) big.push(clusters[i]);
  for (i = 0; i < big.length; i++) { var r = W(big[i]) / H(big[i]); (r < 0.8 ? labels : logos).push(big[i]); }
  if (labels.length !== 7 || logos.length !== 2) {
    alert('Expected 7 labels and 2 logos, found ' + labels.length + ' labels and ' + logos.length + ' logos. Nothing was changed.');
    return;
  }
  labels.sort(function (p, q) { return cy(q) - cy(p); });        // top row first
  var fronts = labels.slice(0, 3).sort(function (p, q) { return cx(p) - cx(q); });
  var dragon = labels[3];
  var backs = labels.slice(4, 7).sort(function (p, q) { return cx(p) - cx(q); });
  logos.sort(function (p, q) { return cy(q) - cy(p); });
  var flameLogo = logos[0], wordmark = logos[1];

  // ---------- 2. helpers
  var lay = doc.layers.add(); lay.name = 'Presentation layout';
  var ab0 = doc.artboards[0].artboardRect;
  var ABW = 1224, ABH = 792, GAP = 120;
  var col = 0, row = 0, PER_ROW = 3;
  function newBoard(name) {
    var L = ab0[0] + col * (ABW + GAP), T = ab0[3] - GAP - row * (ABH + GAP);
    var abd = doc.artboards.add([L, T, L + ABW, T - ABH]); abd.name = name;
    col++; if (col >= PER_ROW) { col = 0; row++; }
    return { L: L, T: T };
  }
  function dup(cluster) {
    var g = lay.groupItems.add();
    for (var k = cluster.items.length - 1; k >= 0; k--) cluster.items[k].duplicate(g, ElementPlacement.PLACEATBEGINNING);
    return g;
  }
  function fitH(g, h) { var b = g.visibleBounds; var s = h / (b[1] - b[3]) * 100; g.resize(s, s, true, true, true, true, s, Transformation.TOPLEFT); }
  function fitBox(g, w, h) { var b = g.visibleBounds; var s = Math.min(w / (b[2] - b[0]), h / (b[1] - b[3])) * 100; g.resize(s, s, true, true, true, true, s, Transformation.TOPLEFT); }
  function place(g, L, T) { var b = g.visibleBounds; g.translate(L - b[0], T - b[1]); }
  function text(s, L, T, size, color, bold) {
    var t = lay.textFrames.add(); t.contents = s; t.left = L; t.top = T;
    var ca = t.textRange.characterAttributes; ca.size = size; ca.fillColor = color || INK;
    try { ca.textFont = app.textFonts.getByName(bold ? 'Helvetica-Bold' : 'Helvetica'); } catch (e) {}
    return t;
  }
  function rect(L, T, w, h, color, stroke) {
    var p = lay.pathItems.rectangle(T, L, w, h);
    if (color) { p.filled = true; p.fillColor = color; } else p.filled = false;
    if (stroke) { p.stroked = true; p.strokeColor = stroke; p.strokeWidth = 0.75; p.strokeDashes = [4, 3]; } else p.stroked = false;
    return p;
  }
  function recolor(obj, c) {
    var t = obj.typename;
    if (t === 'PathItem') { if (obj.filled) obj.fillColor = c; if (obj.stroked) obj.strokeColor = c; }
    else if (t === 'CompoundPathItem') { for (var k = 0; k < obj.pathItems.length; k++) recolor(obj.pathItems[k], c); }
    else if (t === 'GroupItem') { for (var k2 = 0; k2 < obj.pageItems.length; k2++) recolor(obj.pageItems[k2], c); }
    else if (t === 'TextFrame') { obj.textRange.characterAttributes.fillColor = c; }
  }

  // label height in points = scale reference for the sleeve proposal
  var labelH = H(fronts[0]);
  // sleeve flat size proposal derived from the 5 oz woozy mockup (neck circumference ~2.87 vs label height 2.95; sleeve height ~1.55)
  var SLW = labelH * 2.87 / 2.95, SLH = labelH * 1.55 / 2.95;

  function makeSleeve(flav, L, T, w, h) {
    var s = w / SLW;
    var g = lay.groupItems.add(); g.name = 'Neck sleeve – ' + flav.name;
    var bg = g.pathItems.rectangle(T, L, w, h); bg.filled = true; bg.fillColor = flav.col; bg.stroked = false;
    var mark = dup(wordmark); recolor(mark, CREAM);
    var mw = w / 6.2; fitBox(mark, mw, mw);
    var mb = mark.visibleBounds, markW = mb[2] - mb[0], markH = mb[1] - mb[3];
    var stepX = w / 6, stepY = markH * 1.35, rr = 0;
    for (var y = T - markH * 0.15; y > T - h - markH; y -= stepY) {
      var off = (rr % 2) ? stepX / 2 : 0;
      for (var x = L - stepX + off; x < L + w; x += stepX) {
        var m = mark.duplicate(g, ElementPlacement.PLACEATEND); place(m, x + (stepX - markW) / 2, y);
      }
      rr++;
    }
    mark.remove();
    // perforation line
    var perf = g.pathItems.add(); perf.setEntirePath([[L + w * 0.75, T], [L + w * 0.75, T - h]]);
    perf.filled = false; perf.stroked = true; perf.strokeWidth = 1.2 * s; perf.strokeDashes = [3 * s, 4 * s];
    perf.strokeColor = rgb(Math.round(flav.col.red * 0.8), Math.round(flav.col.green * 0.8), Math.round(flav.col.blue * 0.8));
    var clip = g.pathItems.rectangle(T, L, w, h); clip.move(g, ElementPlacement.PLACEATBEGINNING); clip.clipping = true;
    g.clipped = true;
    return g;
  }

  // ---------- 3. boards
  var M = 72, TOP = 120;
  // A. brand marks
  var p = newBoard('01 Brand marks');
  text('BURNING FLAME — BRAND MARKS', p.L + M, p.T - 48, 20, INK, true);
  var g1 = dup(flameLogo); fitBox(g1, 520, 520); place(g1, p.L + M + 20, p.T - TOP - 20);
  var g2 = dup(wordmark); fitBox(g2, 440, 440); place(g2, p.L + 680, p.T - TOP - 60);
  text('Primary · flame lockup', p.L + M + 20, p.T - ABH + 60, 11);
  text('Secondary · wordmark', p.L + 680, p.T - ABH + 60, 11);

  // B. the range
  p = newBoard('02 Front labels');
  text('FRONT LABELS — THE RANGE', p.L + M, p.T - 48, 20, INK, true);
  var LH = 560, x = p.L + M + 40;
  for (i = 0; i < 3; i++) { var f = dup(fronts[i]); fitH(f, LH); place(f, x, p.T - TOP); x += (f.visibleBounds[2] - f.visibleBounds[0]) + 60; }

  // C. one board per flavour: front, back, neck sleeve
  for (i = 0; i < 3; i++) {
    var fl = FLAV[i];
    p = newBoard('0' + (3 + i) + ' ' + fl.name);
    text(fl.name, p.L + M, p.T - 48, 26, fl.col, true);
    text('Heat ' + fl.heat + '  ·  5 FL OZ (148 mL)', p.L + M, p.T - 84, 12);
    var lh = 500;
    var fr = dup(fronts[i]); fitH(fr, lh); place(fr, p.L + M, p.T - TOP - 30);
    var fw = fr.visibleBounds[2] - fr.visibleBounds[0];
    var bk = dup(backs[i]); fitH(bk, lh); place(bk, p.L + M + fw + 48, p.T - TOP - 30);
    var bw = bk.visibleBounds[2] - bk.visibleBounds[0];
    var sx = p.L + M + fw + bw + 110, sw = p.L + ABW - M - sx, sh = sw * SLH / SLW;
    makeSleeve(fl, sx, p.T - TOP - 30, sw, sh);
    text('Front', p.L + M, p.T - TOP - 30 - lh - 14, 10);
    text('Back', p.L + M + fw + 48, p.T - TOP - 30 - lh - 14, 10);
    text('Neck shrink sleeve (concept) — flat, wraps the neck', sx, p.T - TOP - 30 - sh - 14, 10);
    text('Proposed flat ~' + Math.round(SLW) + ' × ' + Math.round(SLH) + ' pt at label scale. Confirm against the bottle supplier\'s sleeve dieline.', sx, p.T - TOP - 30 - sh - 32, 9);
  }

  // D. extra hot routes
  p = newBoard('06 Extra Hot routes');
  text('EXTRA HOT — TWO ROUTES', p.L + M, p.T - 48, 20, INK, true);
  var e1 = dup(fronts[2]); fitH(e1, 540); place(e1, p.L + 260, p.T - TOP);
  var e2 = dup(dragon); fitH(e2, 540); place(e2, p.L + 660, p.T - TOP);
  text('A · Devil', p.L + 260, p.T - TOP - 556, 11, INK, true);
  text('B · Fire-breathing seahorse', p.L + 660, p.T - TOP - 556, 11, INK, true);

  // E. sleeves together
  p = newBoard('07 Neck sleeves');
  text('NECK SHRINK SLEEVES — CONCEPT', p.L + M, p.T - 48, 20, INK, true);
  var sW = ABW - 2 * M, sH = sW * SLH / SLW; if (sH > 190) { sH = 190; sW = sH * SLW / SLH; }
  for (i = 0; i < 3; i++) {
    var yT = p.T - TOP - i * (sH + 40);
    makeSleeve(FLAV[i], p.L + M, yT, sW, sH);
    text(FLAV[i].name, p.L + M + sW + 14, yT - 6, 11, FLAV[i].col, true);
  }

  // ---------- 4. save as NEW files (original stays as it is on disk)
  var base = Folder.desktop + '/BURNING FLAMES_presentation';
  var first = doc.artboards.length - 7 + 1, last = doc.artboards.length;
  var pdf = new PDFSaveOptions(); pdf.compatibility = PDFCompatibility.ACROBAT7; pdf.preserveEditability = false;
  pdf.artboardRange = first + '-' + last; pdf.viewAfterSaving = false;
  doc.saveAs(new File(base + '.pdf'), pdf);
  var ai = new IllustratorSaveOptions(); ai.pdfCompatible = true;
  doc.saveAs(new File(base + '.ai'), ai);
  alert('Done: 7 new artboards on layer "Presentation layout".\nSaved BURNING FLAMES_presentation.ai and .pdf to the Desktop.');
})();
