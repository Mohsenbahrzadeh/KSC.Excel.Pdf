/* KSC.Excel designer - grid: rendering, selection, in-cell editing, clipboard, resizing, drag & drop. */
(function () {
'use strict';
const KXL = window.KXL, S = KXL.S, T = KXL.T, $ = KXL.$, $$ = KXL.$$, el = KXL.el, esc = KXL.esc;
const G = KXL.grid = {};

let wrap, tbl, overlay, actBox, editor, ctx;
let tdMap = new Map(), colTh = [], rowTh = [];
const PXW = (c) => Math.round(c * 7 + 5), PXH = (pt) => Math.round(pt * 4 / 3);
const BAND_COLORS = ['#2e75b6', '#c55a11', '#548235', '#7030a0', '#bf9000', '#c00000'];

G.gridSize = function (sh) {
  const u = KXL.maxUsed(sh);
  sh._gr = Math.max(sh._gr || 0, 60, u.r + 25, S.sel.r2 + 5);
  sh._gc = Math.max(sh._gc || 0, 14, u.c + 4, S.sel.c2 + 2);
};

/* ------------------------------------------------------------------ rendering */
function bandMap(sh) {
  const m = {};
  sh.bands.forEach((b, i) => { for (let r = b.firstRow; r <= b.lastRow; r++) (m[r] = m[r] || []).push(i); });
  for (const r in m) m[r].sort((a, b) => (sh.bands[b].lastRow - sh.bands[b].firstRow) - (sh.bands[a].lastRow - sh.bands[a].firstRow)); // outer first
  return m;
}

function borderCss(def, fallback) {
  if (!def || !def.style) return fallback;
  const w = def.style === 'medium' ? 2 : def.style === 'thick' ? 3 : def.style === 'double' ? 3 : 1;
  const st = (def.style === 'dashed' || def.style === 'mediumdashed') ? 'dashed' : (def.style === 'dotted' || def.style === 'hair') ? 'dotted' : def.style === 'double' ? 'double' : 'solid';
  return w + 'px ' + st + ' ' + (def.color || '#000');
}

function cellCss(sh, st, r, c, rtl) {
  const grid = sh.showGridLines ? '#dfe3e8' : 'transparent';
  let css = '';
  if (st.font) css += "font-family:'" + st.font.replace(/'/g, '') + "',Tahoma,sans-serif;";
  if (st.size) css += 'font-size:' + st.size + 'pt;';
  if (st.bold) css += 'font-weight:700;';
  if (st.italic) css += 'font-style:italic;';
  if (st.underline || st.strike) css += 'text-decoration:' + (st.underline ? 'underline ' : '') + (st.strike ? 'line-through' : '') + ';';
  if (st.color) css += 'color:' + st.color + ';';
  if (st.fill) css += 'background-color:' + st.fill + ';';
  const h = st.h; if (h === 'left' || h === 'right' || h === 'center') css += 'text-align:' + h + ';'; else if (h === 'justify' || h === 'distributed') css += 'text-align:justify;';
  const v = st.v; css += 'vertical-align:' + (v === 'top' ? 'top' : (v === 'center' || v === 'middle') ? 'middle' : 'bottom') + ';';
  if (st.wrap) css += 'white-space:pre-wrap;word-break:break-word;';
  if (st.indent) css += 'padding-' + (rtl ? 'right' : 'left') + ':' + (4 + st.indent * 9) + 'px;';
  if (st.rtl === true) css += 'direction:rtl;'; else if (st.rtl === false) css += 'direction:ltr;';
  // borders (physical sides): a cell draws its own right/bottom; its left/top only when the neighbour has no border there
  const nLs = KXL.styleOf(KXL.cell(r, rtl ? c + 1 : c - 1, sh)), nTs = KXL.styleOf(KXL.cell(r - 1, c, sh));
  const gridB = '1px solid ' + grid, none = '1px solid transparent';
  css += 'border-right:' + borderCss(st.right, gridB) + ';';
  css += 'border-bottom:' + borderCss(st.bottom, gridB) + ';';
  css += 'border-left:' + (st.left && !nLs.right ? borderCss(st.left, gridB) : none) + ';';
  css += 'border-top:' + (st.top && !nTs.bottom ? borderCss(st.top, gridB) : none) + ';';
  return css;
}

function display(cell) {
  if (!cell || cell.v == null || cell.v === '') return '';
  const v = String(cell.v);
  if (cell.bc) return '<span class="ph" title="barcode">\u25A6 ' + esc(cell.bc.split(':')[0].toUpperCase()) + ' ' + esc(v) + '</span>';
  if (v[0] === '=') return '<span class="fx">' + esc(v) + '</span>';
  return esc(v).replace(/\{\{(.*?)\}\}/g, '<span class="ph">{{$1}}</span>');
}

G.render = function () {
  const sh = KXL.cur(); if (!sh) return;
  G.gridSize(sh);
  const rtl = !!sh.rtl; const R = sh._gr, C = sh._gc;
  const bm = bandMap(sh);
  const anchors = new Map(), covered = new Set();
  for (const m of sh.merges) {
    const g = KXL.parseRange(m); if (!g) continue;
    anchors.set(g.r1 + ',' + g.c1, g);
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (r !== g.r1 || c !== g.c1) covered.add(r + ',' + c);
  }
  const frR = sh.freezeRows || 0, frC = sh.freezeCols || 0;
  const SECT = { reportHeader: ['#546e7a', 'secReportHeader'], pageHeader: ['#00796b', 'secPageHeader'], reportFooter: ['#6d4c41', 'secReportFooter'] };
  const sectAt = {}; (sh.sections || []).forEach(x => { for (let r = x.firstRow; r <= x.lastRow; r++) sectAt[r] = x; });
  const colBand = {}; sh.bands.forEach((b, i) => { if (b.orientation === 'columns') for (let c = b.firstCol; c <= b.lastCol; c++) colBand[c] = i; });
  const parts = [];
  parts.push('<table class="grid' + (rtl ? ' rtl' : '') + '" dir="' + (rtl ? 'rtl' : 'ltr') + '"><colgroup><col style="width:46px">');
  for (let c = 1; c <= C; c++) parts.push('<col style="width:' + PXW(KXL.colWidth(c, sh)) + 'px">');
  parts.push('</colgroup><thead><tr><th class="corner" data-act="all"></th>');
  for (let c = 1; c <= C; c++) parts.push('<th class="ch' + (colBand[c] !== undefined ? ' cband b' + (colBand[c] % 6) : '') + '" data-c="' + c + '" title="' + (colBand[c] !== undefined ? esc(sh.bands[colBand[c]].name || sh.bands[colBand[c]].source) : '') + '">' + KXL.colName(c) + '<i class="grip" data-grip="c" data-i="' + c + '"></i></th>');
  parts.push('</tr></thead><tbody>');
  for (let r = 1; r <= R; r++) {
    const bi = bm[r];
    let tag = '', bcls = '';
    if (bi) {
      bcls = ' band b' + (bi[bi.length - 1] % 6);
      for (const i of bi) { const b = sh.bands[i]; if (b.firstRow === r) tag += '<em class="btag" style="background:' + BAND_COLORS[i % 6] + '" title="' + esc(b.source) + '">' + esc(b.name || b.source) + '</em>'; }
    }
    const sx = sectAt[r]; if (sx && sx.firstRow === r) { const k = SECT[sx.kind] || ['#777', null]; tag += '<em class="btag stag" style="background:' + k[0] + '">' + esc(sx.name || (k[1] ? T(k[1]) : sx.kind)) + '</em>'; }
    parts.push('<tr style="height:' + PXH(KXL.rowHeight(r, sh)) + 'px"><th class="rh' + bcls + (sx ? ' sect' : '') + '" data-r="' + r + '"' + (sx ? ' style="--sc:' + (SECT[sx.kind] || ['#777'])[0] + '"' : '') + '>' + r + tag + '<i class="grip" data-grip="r" data-i="' + r + '"></i></th>');
    for (let c = 1; c <= C; c++) {
      const k = r + ',' + c;
      if (covered.has(k)) continue;
      const cell = sh._m.get(k); const st = KXL.styleOf(cell);
      let span = '';
      const g = anchors.get(k);
      if (g) { if (g.c2 > g.c1) span += ' colspan="' + (g.c2 - g.c1 + 1) + '"'; if (g.r2 > g.r1) span += ' rowspan="' + (g.r2 - g.r1 + 1) + '"'; }
      let cls = bi ? 'inband b' + (bi[bi.length - 1] % 6) : '';
      if (r === frR) cls += ' fb'; if (c === frC) cls += ' fe';
      if (cell && cell.l) cls += ' lnk';
      const numeric = cell && cell.v && !st.h && /^-?\d+(\.\d+)?$/.test(cell.v) && cell.t !== 'text';
      parts.push('<td data-r="' + r + '" data-c="' + c + '" class="' + cls.trim() + '"' + span + ' style="' + cellCss(sh, st, r, c, rtl) + (numeric ? 'text-align:right;' : '') + '">' + display(cell) + '</td>');
    }
    parts.push('</tr>');
  }
  parts.push('</tbody></table>');
  tbl = null;
  // keep overlay/editor nodes, replace table
  const old = $('table.grid', wrap); if (old) old.remove();
  wrap.insertAdjacentHTML('afterbegin', parts.join(''));
  tbl = $('table.grid', wrap);
  tdMap = new Map(); $$('td[data-r]', tbl).forEach(td => tdMap.set(td.dataset.r + ',' + td.dataset.c, td));
  colTh = $$('th.ch', tbl); rowTh = $$('th.rh', tbl);
  $$('.gimg', wrap).forEach(x => x.remove());
  renderImages(sh);
  renderCharts(sh);
  G.paint();
};

function renderCharts(sh) {
  for (const ch of sh.charts || []) {
    const td = tdMap.get(ch.row + ',' + ch.col) || tdMap.get(ch.row + ',' + (ch.col + 1)); if (!td) continue;
    const w = ch.width || 480, h = ch.height || 288; const rtl = !!sh.rtl;
    const box = el('div', { class: 'gimg gchart', title: ch.name || '' }, '\u{1F4CA} ' + (ch.title || ch.name || 'Chart') + ' (' + (ch.type || 'column') + ')');
    box.style.cssText = 'position:absolute;top:' + td.offsetTop + 'px;left:' + (rtl ? td.offsetLeft + td.offsetWidth - w : td.offsetLeft) + 'px;width:' + w + 'px;height:' + h + 'px;pointer-events:none;border:2px dashed #2e75b6;background:rgba(46,117,182,.07);color:#2e75b6;display:flex;align-items:center;justify-content:center;font-size:13px;border-radius:6px;z-index:1';
    wrap.append(box);
  }
}

function renderImages(sh) {
  for (const im of sh.images) {
    const data = S.blobs[im.dataRef]; if (!data) continue;
    const td = tdMap.get(im.row + ',' + im.col) || tdMap.get(im.row + ',' + (im.col + 1)); if (!td) continue;
    const rtl = !!sh.rtl;
    const img = el('img', { class: 'gimg', src: 'data:image/' + (im.ext || 'png') + ';base64,' + data, draggable: 'false', title: im.name || '' });
    img.style.cssText = 'position:absolute;top:' + td.offsetTop + 'px;left:' + (rtl ? td.offsetLeft + td.offsetWidth - im.width : td.offsetLeft) + 'px;width:' + im.width + 'px;height:' + im.height + 'px;opacity:.9;pointer-events:none';
    wrap.append(img);
  }
}

/* ------------------------------------------------------------------ selection */
function rectOf(r, c) {
  const td = tdMap.get(r + ',' + c) || anchorTd(r, c);
  return td ? { l: td.offsetLeft, t: td.offsetTop, w: td.offsetWidth, h: td.offsetHeight } : null;
}
function anchorTd(r, c) { const g = KXL.mergeAt(r, c); return g ? tdMap.get(g.r1 + ',' + g.c1) : null; }

G.expandToMerges = function (s) {
  let ch = true;
  while (ch) {
    ch = false;
    for (const m of KXL.cur().merges) {
      const g = KXL.parseRange(m); if (!g) continue;
      if (g.r1 <= s.r2 && g.r2 >= s.r1 && g.c1 <= s.c2 && g.c2 >= s.c1) {
        if (g.r1 < s.r1) { s.r1 = g.r1; ch = true; } if (g.r2 > s.r2) { s.r2 = g.r2; ch = true; }
        if (g.c1 < s.c1) { s.c1 = g.c1; ch = true; } if (g.c2 > s.c2) { s.c2 = g.c2; ch = true; }
      }
    }
  }
  return s;
};

G.select = function (ar, ac, fr, fc, keepAnchor) {
  const sh = KXL.cur();
  fr = fr == null ? ar : fr; fc = fc == null ? ac : fc;
  ar = Math.max(1, Math.min(ar, 1048576)); ac = Math.max(1, Math.min(ac, 16384)); fr = Math.max(1, fr); fc = Math.max(1, fc);
  const s = { r1: Math.min(ar, fr), r2: Math.max(ar, fr), c1: Math.min(ac, fc), c2: Math.max(ac, fc), ar, ac };
  G.expandToMerges(s);
  const need = (s.r2 > sh._gr - 3) || (s.c2 > sh._gc - 1);
  S.sel = s;
  if (need) { sh._gr = Math.max(sh._gr, s.r2 + 20); sh._gc = Math.max(sh._gc, s.c2 + 3); G.render(); } else G.paint();
  KXL.fn.selectionChanged && KXL.fn.selectionChanged();
};

G.paint = function () {
  if (!tbl) return;
  const s = S.sel;
  const a = rectOf(s.r1, s.c1), b = rectOf(s.r2, s.c2);
  if (a && b) {
    const l = Math.min(a.l, b.l), t = Math.min(a.t, b.t), r = Math.max(a.l + a.w, b.l + b.w), bt = Math.max(a.t + a.h, b.t + b.h);
    overlay.style.cssText = 'display:block;left:' + (l - 1) + 'px;top:' + (t - 1) + 'px;width:' + (r - l + 1) + 'px;height:' + (bt - t + 1) + 'px';
  }
  const ac = rectOf(s.ar, s.ac);
  if (ac) actBox.style.cssText = 'display:block;left:' + (ac.l - 1) + 'px;top:' + (ac.t - 1) + 'px;width:' + (ac.w + 1) + 'px;height:' + (ac.h + 1) + 'px';
  colTh.forEach((th, i) => th.classList.toggle('hl', i + 1 >= s.c1 && i + 1 <= s.c2));
  rowTh.forEach((th, i) => th.classList.toggle('hl', i + 1 >= s.r1 && i + 1 <= s.r2));
};

G.scrollIntoView = function () {
  const td = tdMap.get(S.sel.ar + ',' + S.sel.ac) || anchorTd(S.sel.ar, S.sel.ac); if (!td) return;
  const sc = $('#gridscroll'); const pad = 60;
  const l = td.offsetLeft, t = td.offsetTop, w = td.offsetWidth, h = td.offsetHeight;
  if (t < sc.scrollTop + 26) sc.scrollTop = Math.max(0, t - 26); else if (t + h > sc.scrollTop + sc.clientHeight - 12) sc.scrollTop = t + h - sc.clientHeight + 16;
  if (l < sc.scrollLeft + 46) sc.scrollLeft = Math.max(0, l - 46 - pad); else if (l + w > sc.scrollLeft + sc.clientWidth - 12) sc.scrollLeft = l + w - sc.clientWidth + 16;
};

/* ------------------------------------------------------------------ editing */
G.startEdit = function (initial) {
  if (S.readOnlyTemplate) return;
  const s = S.sel; const sh = KXL.cur();
  const td = tdMap.get(s.ar + ',' + s.ac) || anchorTd(s.ar, s.ac); if (!td) return;
  const c = KXL.cell(s.ar, s.ac);
  S.editing = { r: s.ar, c: s.ac };
  editor.style.cssText = 'display:block;left:' + (td.offsetLeft - 1) + 'px;top:' + (td.offsetTop - 1) + 'px;min-width:' + (td.offsetWidth + 1) + 'px;min-height:' + (td.offsetHeight + 1) + 'px;' + cellCss(sh, KXL.styleOf(c), s.ar, s.ac, !!sh.rtl).replace(/border-[a-z]+:[^;]+;/g, '') + 'direction:' + (sh.rtl ? 'rtl' : 'ltr') + ';';
  editor.value = initial != null ? initial : (c && c.v ? c.v : '');
  editor.focus();
  if (initial == null) editor.select(); else editor.setSelectionRange(editor.value.length, editor.value.length);
  sizeEditor();
};
function sizeEditor() { editor.style.height = 'auto'; editor.style.height = Math.max(editor.scrollHeight, parseInt(editor.style.minHeight) || 0) + 'px'; editor.style.width = Math.max(editor.scrollWidth + 6, parseInt(editor.style.minWidth) || 0) + 'px'; }
G.commitEdit = function (move) {
  if (!S.editing) return;
  const e = S.editing; S.editing = null;
  const v = editor.value; editor.style.display = 'none';
  const cell = KXL.cell(e.r, e.c); const old = cell && cell.v ? cell.v : '';
  if (v !== old) KXL.mutate(() => KXL.setValue(e.r, e.c, v));
  G.render();
  if (move) moveSel(move.dr, move.dc, false); else G.paint();
  $('#gridwrap').focus({ preventScroll: true });
};
G.cancelEdit = function () { if (!S.editing) return; S.editing = null; editor.style.display = 'none'; wrap.focus({ preventScroll: true }); G.paint(); };

function moveSel(dr, dc, extend) {
  const s = S.sel; const rtl = !!KXL.cur().rtl;
  if (extend) {
    const fr = (s.ar === s.r1 ? s.r2 : s.r1) + dr, fc = (s.ac === s.c1 ? s.c2 : s.c1) + dc;
    G.select(s.ar, s.ac, Math.max(1, fr), Math.max(1, fc));
  } else {
    // keep the cursor inside a multi-cell selection when pressing Enter/Tab
    G.select(Math.max(1, s.ar + dr), Math.max(1, s.ac + dc));
  }
  G.scrollIntoView();
}

/* ------------------------------------------------------------------ styling helpers */
G.eachCell = function (fn, limit) {
  const s = S.sel; const sh = KXL.cur();
  const r2 = Math.min(s.r2, sh._gr), c2 = Math.min(s.c2, sh._gc);
  let n = 0;
  for (let r = s.r1; r <= r2; r++) for (let c = s.c1; c <= c2; c++) { if (++n > (limit || 20000)) return; fn(r, c); }
};
G.applyStyle = function (patch) {
  KXL.mutate(() => G.eachCell((r, c) => {
    const cell = KXL.ensureCell(r, c); const st = Object.assign({}, KXL.styleOf(cell), typeof patch === 'function' ? patch(KXL.styleOf(cell), r, c) : patch);
    cell.s = KXL.styleId(st);
    if (!cell.s && !cell.v && !cell.l) KXL.cur()._m.delete(r + ',' + c);
  }));
  G.render();
};
G.setBorder = function (mode) {
  const s = S.sel; const b = S.border; const def = { style: b.style, color: b.color };
  const sides = { left: 'left', right: 'right', top: 'top', bottom: 'bottom' };
  const opp = { left: 'right', right: 'left', top: 'bottom', bottom: 'top' };
  const rtl = !!KXL.cur().rtl;
  KXL.mutate(() => {
    const setEdge = (r, c, side, d) => {
      if (r < 1 || c < 1) return;
      const cell = KXL.ensureCell(r, c); const st = Object.assign({}, KXL.styleOf(cell));
      if (d) st[side] = Object.assign({}, d); else delete st[side];
      cell.s = KXL.styleId(st);
    };
    const edge = (r, c, side, d) => {
      setEdge(r, c, side, d);
      const dr = side === 'top' ? -1 : side === 'bottom' ? 1 : 0, dc = (side === 'left' ? -1 : side === 'right' ? 1 : 0) * (rtl ? -1 : 1);
      if (r + dr >= 1 && c + dc >= 1) setEdge(r + dr, c + dc, opp[side], d);
    };
    const clear = mode === 'none';
    const d = clear ? null : def;
    G.eachCell((r, c) => {
      const top = r === s.r1, bot = r === s.r2, lef = rtl ? c === s.c2 : c === s.c1, rig = rtl ? c === s.c1 : c === s.c2;
      const all = mode === 'all' || mode === 'none';
      if (all || (mode === 'outer' || mode === 'thick') && top || mode === 'top' && top) edge(r, c, 'top', mode === 'thick' ? { style: 'thick', color: def.color } : d);
      if (all || (mode === 'outer' || mode === 'thick') && bot || mode === 'bottom' && bot || mode === 'double' && bot) edge(r, c, 'bottom', mode === 'thick' ? { style: 'thick', color: def.color } : mode === 'double' ? { style: 'double', color: def.color } : d);
      if (all || (mode === 'outer' || mode === 'thick') && lef || mode === 'left' && lef) edge(r, c, 'left', mode === 'thick' ? { style: 'thick', color: def.color } : d);
      if (all || (mode === 'outer' || mode === 'thick') && rig || mode === 'right' && rig) edge(r, c, 'right', mode === 'thick' ? { style: 'thick', color: def.color } : d);
    });
    // tidy: remove empty cells created only to hold nothing
    for (const [k, cell] of KXL.cur()._m) if (!cell.v && !cell.s && !cell.l) KXL.cur()._m.delete(k);
  });
  G.render();
};

G.merge = function () {
  const s = S.sel; if (s.r1 === s.r2 && s.c1 === s.c2) return;
  KXL.mutate(() => {
    const sh = KXL.cur();
    sh.merges = sh.merges.filter(m => { const g = KXL.parseRange(m); return !(g.r1 <= s.r2 && g.r2 >= s.r1 && g.c1 <= s.c2 && g.c2 >= s.c1); });
    const tl = KXL.ensureCell(s.r1, s.c1); const st = KXL.styleOf(tl);
    G.eachCell((r, c) => { if (r === s.r1 && c === s.c1) return; const x = KXL.ensureCell(r, c); delete x.v; delete x.t; x.s = KXL.styleId(st); });
    sh.merges.push(KXL.rangeStr(s));
  });
  G.render();
};
G.unmerge = function () {
  const s = S.sel;
  KXL.mutate(() => { const sh = KXL.cur(); sh.merges = sh.merges.filter(m => { const g = KXL.parseRange(m); return !(g.r1 <= s.r2 && g.r2 >= s.r1 && g.c1 <= s.c2 && g.c2 >= s.c1); }); });
  G.render();
};
G.clearContents = function () {
  KXL.mutate(() => G.eachCell((r, c) => { const x = KXL.cell(r, c); if (x) { delete x.v; delete x.t; delete x.l; delete x.bc; if (!x.s) KXL.cur()._m.delete(r + ',' + c); } }));
  G.render();
};
G.clearFormats = function () {
  KXL.mutate(() => G.eachCell((r, c) => { const x = KXL.cell(r, c); if (x) { delete x.s; if (!x.v && !x.l) KXL.cur()._m.delete(r + ',' + c); } }));
  G.render();
};

/* ------------------------------------------------------------------ clipboard */
function selText() {
  const s = S.sel; const rows = [];
  for (let r = s.r1; r <= Math.min(s.r2, s.r1 + 2000); r++) { const cols = []; for (let c = s.c1; c <= Math.min(s.c2, s.c1 + 200); c++) { const x = KXL.cell(r, c); cols.push(x && x.v ? String(x.v).replace(/\t|\r?\n/g, ' ') : ''); } rows.push(cols.join('\t')); }
  return rows.join('\r\n');
}
G.copy = function (e, cut) {
  const s = S.sel; const sh = KXL.cur(); const items = [];
  G.eachCell((r, c) => { const x = KXL.cell(r, c); if (x) items.push({ dr: r - s.r1, dc: c - s.c1, v: x.v, s: x.s ? JSON.parse(JSON.stringify(KXL.styleOf(x))) : null, t: x.t, l: x.l, bc: x.bc }); });
  const merges = sh.merges.map(m => KXL.parseRange(m)).filter(g => g && g.r1 >= s.r1 && g.r2 <= s.r2 && g.c1 >= s.c1 && g.c2 <= s.c2).map(g => ({ dr: g.r1 - s.r1, dc: g.c1 - s.c1, h: g.r2 - g.r1, w: g.c2 - g.c1 }));
  const text = selText();
  S.clip = { items, merges, text, rows: s.r2 - s.r1 + 1, cols: s.c2 - s.c1 + 1 };
  if (e && e.clipboardData) { e.clipboardData.setData('text/plain', text); e.preventDefault(); }
  else if (navigator.clipboard) navigator.clipboard.writeText(text).catch(() => { });
  if (cut) G.clearContents();
};
G.paste = function (text) {
  const s = S.sel; const sh = KXL.cur();
  KXL.mutate(() => {
    if (S.clip && text != null && S.clip.text.replace(/\r/g, '') === String(text).replace(/\r/g, '')) {
      for (let dr = 0; dr < S.clip.rows; dr++) for (let dc = 0; dc < S.clip.cols; dc++) { const x = KXL.cell(s.r1 + dr, s.c1 + dc); if (x) { delete x.v; delete x.t; delete x.l; delete x.s; delete x.bc; } }
      for (const it of S.clip.items) {
        const r = s.r1 + it.dr, c = s.c1 + it.dc; const x = KXL.ensureCell(r, c);
        if (it.v) x.v = it.v; if (it.t) x.t = it.t; if (it.l) x.l = it.l; if (it.bc) x.bc = it.bc; x.s = it.s ? KXL.styleId(it.s) : undefined; if (!x.s) delete x.s;
      }
      for (const m of S.clip.merges) { const g = { r1: s.r1 + m.dr, c1: s.c1 + m.dc, r2: s.r1 + m.dr + m.h, c2: s.c1 + m.dc + m.w }; sh.merges = sh.merges.filter(x => { const q = KXL.parseRange(x); return !(q.r1 <= g.r2 && q.r2 >= g.r1 && q.c1 <= g.c2 && q.c2 >= g.c1); }); sh.merges.push(KXL.rangeStr(g)); }
    } else if (text != null) {
      String(text).replace(/\r/g, '').split('\n').forEach((line, i) => line.split('\t').forEach((v, j) => { if (!(line === '' && i > 0 && v === '')) KXL.setValue(s.r1 + i, s.c1 + j, v); }));
    }
  });
  G.render();
};

/* ------------------------------------------------------------------ field insertion (drag & drop / double click) */
G.absBandSource = function (sh, idx) {   // absolute source of band idx (parents' sources prepended)
  const b = sh.bands[idx]; let src = b.source;
  const parents = sh.bands.filter((p, i) => i !== idx && p.firstRow <= b.firstRow && p.lastRow >= b.lastRow && !(p.firstRow === b.firstRow && p.lastRow === b.lastRow))
    .sort((x, y) => (y.lastRow - y.firstRow) - (x.lastRow - x.firstRow));
  for (const p of parents.reverse()) src = p.source + '.' + src;
  return src;
};

G.insertField = function (info, r, c, mode) {
  // info: {chain:[{name,type}...], leaf:{name,type}}  chain = ancestors (lists/objects) in order
  const sh = KXL.cur();
  const lists = info.chain.filter(x => x.type === 'list');
  const names = info.chain.map(x => x.name);
  let text;
  KXL.mutate(() => {
    let rest = info.path.slice();            // full path segments
    if (lists.length && mode !== 'sum') {
      // bands already covering this row, outer -> inner
      const covering = []; sh.bands.forEach((b, i) => { if (r >= b.firstRow && r <= b.lastRow) covering.push(i); });
      covering.sort((a, b) => (sh.bands[b].lastRow - sh.bands[b].firstRow) - (sh.bands[a].lastRow - sh.bands[a].firstRow));
      const listPaths = []; let acc = [];
      for (const seg of info.chain) { acc.push(seg.name); if (seg.type === 'list') listPaths.push(acc.join('.')); }
      let matched = 0, consumed = 0;
      for (const bi of covering) { const abs = G.absBandSource(sh, bi); if (matched < listPaths.length && abs === listPaths[matched]) { matched++; } }
      if (matched < lists.length) {
        // create a band (for the innermost list) on this row
        const innerIdx = lists.length - 1; const full = listPaths[innerIdx];
        const parentAbs = matched > 0 ? listPaths[matched - 1] : '';
        const source = parentAbs ? full.slice(parentAbs.length + 1) : full;
        const name = uniqueBandName(sh, full.split('.').pop());
        if (!sh.bands.some(b => b.firstRow === r && b.lastRow === r && b.source === source)) sh.bands.push({ name, source, firstRow: r, lastRow: r });
        matched = lists.length;
      }
      // relative path = segments after the last list in the chain
      const lastListPos = info.chain.map(x => x.type).lastIndexOf('list');
      rest = info.path.slice(lastListPos + 1);
    }
    if (mode === 'sum') rest = [info.leaf.name];
    let expr = rest.join('.');
    if (mode === 'jalali') expr += '|fa';
    text = '{{' + expr + '}}';
    if (mode === 'sum') {
      const lastList = info.chain[info.chain.map(x => x.type).lastIndexOf('list')];
      const bn = bandNameForList(sh, info);
      text = '=SUM([' + bn + '.' + info.leaf.name + '])';
    }
    KXL.setValue(r, c, text);
    const x = KXL.cell(r, c);
    if (info.leaf.type === 'date' && mode !== 'jalali' && !KXL.styleOf(x).format) { x.s = KXL.styleId(Object.assign({}, KXL.styleOf(x), { format: 'yyyy/mm/dd' })); }
  });
  G.render(); KXL.fn.changed && KXL.fn.changed();
  return text;
};
function bandNameForList(sh, info) {
  const idx = info.chain.map(x => x.type).lastIndexOf('list');
  const full = info.chain.slice(0, idx + 1).map(x => x.name).join('.');
  for (let i = 0; i < sh.bands.length; i++) if (G.absBandSource(sh, i) === full) return sh.bands[i].name;
  return info.chain[idx].name;
}
function uniqueBandName(sh, base) { let n = base, i = 1; while (sh.bands.some(b => b.name === n)) n = base + (++i); return n; }
G.uniqueBandName = uniqueBandName;

/* ------------------------------------------------------------------ events */
function tdOf(e) { const t = e.target.closest ? e.target.closest('td[data-r]') : null; return t; }

G.init = function () {
  wrap = $('#gridwrap');
  overlay = el('div', { id: 'selbox' }); actBox = el('div', { id: 'actbox' });
  editor = el('textarea', { id: 'cell-editor', spellcheck: 'false', rows: '1' });
  wrap.append(overlay, actBox, editor);
  ctx = $('#ctxmenu');

  let dragging = false, resizing = null;

  wrap.addEventListener('mousedown', (e) => {
    if (e.button === 2) { const t = tdOf(e); if (t) { const r = +t.dataset.r, c = +t.dataset.c; const s = S.sel; if (!(r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2)) G.select(r, c); } return; }
    if (e.target === editor) return;
    if (S.editing) G.commitEdit();
    const grip = e.target.closest && e.target.closest('.grip');
    if (grip) { startResize(e, grip); return; }
    const td = tdOf(e);
    const sh = KXL.cur();
    if (td) {
      const r = +td.dataset.r, c = +td.dataset.c;
      if (e.shiftKey) G.select(S.sel.ar, S.sel.ac, r, c); else G.select(r, c);
      dragging = 'cell'; wrap.focus({ preventScroll: true }); e.preventDefault(); return;
    }
    const ch = e.target.closest && e.target.closest('th.ch');
    if (ch) { const c = +ch.dataset.c; if (e.shiftKey) G.select(1, S.sel.ac, sh._gr, c); else G.select(1, c, sh._gr, c); dragging = 'col'; wrap.focus({ preventScroll: true }); e.preventDefault(); return; }
    const rh = e.target.closest && e.target.closest('th.rh');
    if (rh) { const r = +rh.dataset.r; if (e.shiftKey) G.select(S.sel.ar, 1, r, sh._gc); else G.select(r, 1, r, sh._gc); dragging = 'row'; wrap.focus({ preventScroll: true }); e.preventDefault(); return; }
    if (e.target.closest && e.target.closest('th.corner')) { G.select(1, 1, sh._gr, sh._gc); wrap.focus({ preventScroll: true }); }
  });
  wrap.addEventListener('mouseover', (e) => {
    if (!dragging || resizing) return;
    const sh = KXL.cur();
    if (dragging === 'cell') { const td = tdOf(e); if (td) G.select(S.sel.ar, S.sel.ac, +td.dataset.r, +td.dataset.c); }
    else if (dragging === 'col') { const ch = e.target.closest('th.ch'); if (ch) G.select(1, S.sel.ac, sh._gr, +ch.dataset.c); }
    else if (dragging === 'row') { const rh = e.target.closest('th.rh'); if (rh) G.select(S.sel.ar, 1, +rh.dataset.r, sh._gc); }
  });
  document.addEventListener('mouseup', () => { dragging = false; });
  wrap.addEventListener('dblclick', (e) => { const td = tdOf(e); if (td) { G.select(+td.dataset.r, +td.dataset.c); G.startEdit(); } });

  /* resize columns / rows */
  function startResize(e, grip) {
    e.preventDefault(); e.stopPropagation();
    const axis = grip.dataset.grip, idx = +grip.dataset.i; const sh = KXL.cur(); const rtl = !!sh.rtl;
    const th = grip.parentElement; const rect = th.getBoundingClientRect();
    const colEl = axis === 'c' ? $$('colgroup col', tbl)[idx] : null; const rowEl = axis === 'r' ? th.parentElement : null;
    resizing = { axis, idx };
    const move = (ev) => {
      if (axis === 'c') { const px = Math.max(12, rtl ? rect.right - ev.clientX : ev.clientX - rect.left); colEl.style.width = px + 'px'; resizing.val = (px - 5) / 7; }
      else { const px = Math.max(8, ev.clientY - rect.top); rowEl.style.height = px + 'px'; resizing.val = px * 3 / 4; }
    };
    const up = () => {
      document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up);
      const v = resizing.val; const rs = resizing; resizing = null;
      if (v) KXL.mutate(() => {
        const targets = [];
        const s = S.sel;
        if (rs.axis === 'c') { if (s.c1 <= rs.idx && rs.idx <= s.c2 && s.r1 === 1 && s.r2 >= sh._gr) for (let c = s.c1; c <= s.c2; c++) targets.push(c); else targets.push(rs.idx); targets.forEach(c => sh._colw.set(c, Math.round(v * 10) / 10)); }
        else { if (s.r1 <= rs.idx && rs.idx <= s.r2 && s.c1 === 1 && s.c2 >= sh._gc) for (let r = s.r1; r <= s.r2; r++) targets.push(r); else targets.push(rs.idx); targets.forEach(r => sh._rowh.set(r, Math.round(v * 2) / 2)); }
      });
      G.render();
    };
    document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
  }

  /* editor */
  editor.addEventListener('keydown', (e) => {
    const rtl = !!KXL.cur().rtl;
    if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); G.commitEdit({ dr: e.shiftKey ? -1 : 1, dc: 0 }); }
    else if (e.key === 'Enter' && e.altKey) { const p = editor.selectionStart; editor.setRangeText('\n', p, editor.selectionEnd, 'end'); sizeEditor(); e.preventDefault(); }
    else if (e.key === 'Tab') { e.preventDefault(); G.commitEdit({ dr: 0, dc: e.shiftKey ? -1 : 1 }); }
    else if (e.key === 'Escape') { e.preventDefault(); G.cancelEdit(); }
    e.stopPropagation();
  });
  editor.addEventListener('input', () => { sizeEditor(); const fb = $('#formula-input'); if (fb) fb.value = editor.value; });
  editor.addEventListener('blur', () => { setTimeout(() => { if (S.editing && document.activeElement !== editor && !(document.activeElement && document.activeElement.id === 'formula-input')) G.commitEdit(); }, 0); });

  /* keyboard on the grid */
  wrap.addEventListener('keydown', (e) => {
    if (S.editing) return;
    const rtl = !!KXL.cur().rtl; const k = e.key; const mod = e.ctrlKey || e.metaKey;
    const arrow = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, rtl ? 1 : -1], ArrowRight: [0, rtl ? -1 : 1] }[k];
    if (arrow) {
      e.preventDefault();
      if (mod) { // jump to data edge
        const sh = KXL.cur(); let r = S.sel.ar, c = S.sel.ac; const u = KXL.maxUsed(sh);
        if (arrow[0]) r = arrow[0] < 0 ? 1 : Math.max(u.r, 1); if (arrow[1]) c = arrow[1] < 0 ? 1 : Math.max(u.c, 1);
        G.select(r, c); G.scrollIntoView(); return;
      }
      moveSel(arrow[0], arrow[1], e.shiftKey); return;
    }
    if (k === 'Tab') { e.preventDefault(); moveSel(0, e.shiftKey ? -1 : 1, false); return; }
    if (k === 'Enter') { e.preventDefault(); G.startEdit(); return; }
    if (k === 'F2') { e.preventDefault(); G.startEdit(); return; }
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); G.clearContents(); return; }
    if (k === 'Home') { e.preventDefault(); G.select(mod ? 1 : S.sel.ar, 1); G.scrollIntoView(); return; }
    if (k === 'PageDown') { e.preventDefault(); moveSel(20, 0, e.shiftKey); return; }
    if (k === 'PageUp') { e.preventDefault(); moveSel(-20, 0, e.shiftKey); return; }
    if (mod) return;                              // shortcuts handled globally
    if (k.length === 1 && !e.altKey) { e.preventDefault(); G.startEdit(k); }
    else if ((k === 'Process' || k === 'Unidentified' || k === 'Dead') && !e.altKey) G.startEdit('');   // IME / dead keys: let the text flow into the editor
  });

  /* clipboard events */
  document.addEventListener('copy', (e) => { if (inGrid()) G.copy(e, false); });
  document.addEventListener('cut', (e) => { if (inGrid()) G.copy(e, true); });
  document.addEventListener('paste', (e) => { if (inGrid()) { e.preventDefault(); G.paste(e.clipboardData.getData('text/plain')); } });
  function inGrid() { const a = document.activeElement; return !S.editing && (a === wrap || wrap.contains(a)); }

  /* drag & drop of fields */
  wrap.addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('text/kxl-field')) { const td = tdOf(e); if (td) { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; } } });
  wrap.addEventListener('drop', (e) => {
    const raw = e.dataTransfer.getData('text/kxl-field'); if (!raw) return;
    const td = tdOf(e); if (!td) return; e.preventDefault();
    const info = JSON.parse(raw); const r = +td.dataset.r, c = +td.dataset.c;
    G.select(r, c); G.insertField(info, r, c, e.shiftKey ? 'jalali' : undefined);
  });

  /* context menu */
  wrap.addEventListener('contextmenu', (e) => {
    const td = tdOf(e), th = e.target.closest && e.target.closest('th.rh,th.ch');
    if (!td && !th) return; e.preventDefault();
    if (th && th.classList.contains('rh')) { const r = +th.dataset.r; if (!(r >= S.sel.r1 && r <= S.sel.r2 && S.sel.c1 === 1)) G.select(r, 1, r, KXL.cur()._gc); }
    if (th && th.classList.contains('ch')) { const c = +th.dataset.c; if (!(c >= S.sel.c1 && c <= S.sel.c2 && S.sel.r1 === 1)) G.select(1, c, KXL.cur()._gr, c); }
    KXL.fn.showContextMenu && KXL.fn.showContextMenu(e.clientX, e.clientY);
  });
  document.addEventListener('mousedown', (e) => { if (ctx && !ctx.contains(e.target)) ctx.style.display = 'none'; });
  window.addEventListener('resize', () => G.paint());
};

})();
