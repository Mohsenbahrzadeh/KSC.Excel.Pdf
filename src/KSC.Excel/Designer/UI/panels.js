/* KSC.Excel designer - ribbon, formula bar, panels, dialogs, file operations. */
(function () {
'use strict';
const KXL = window.KXL, S = KXL.S, T = KXL.T, $ = KXL.$, $$ = KXL.$$, el = KXL.el, esc = KXL.esc, G = KXL.grid;
const UI = KXL.ui;

/* ------------------------------------------------------------------ icons */
const IC = {
  new: '<path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4M12 11v6M9 14h6"/>',
  open: '<path d="M3 7h6l2 2h10v10H3z"/>',
  save: '<path d="M5 3h12l3 3v15H5z"/><path d="M8 3v5h7V3M8 21v-7h8v7"/>',
  eye: '<path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
  dl: '<path d="M12 3v12m0 0l-4-4m4 4l4-4M4 20h16"/>',
  json: '<path d="M8 4c-3 0-3 3-3 4s-1 4-2 4c1 0 2 3 2 4s0 4 3 4M16 4c3 0 3 3 3 4s1 4 2 4c-1 0-2 3-2 4s0 4-3 4"/>',
  undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h10a6 6 0 010 12h-3"/>',
  redo: '<path d="M15 14l5-5-5-5"/><path d="M20 9H10a6 6 0 000 12h3"/>',
  left: '<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>', center: '<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>', right: '<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>',
  top: '<path d="M4 4h16M12 8v12M8 12l4-4 4 4"/>', middle: '<path d="M4 12h16M12 4v6M9 7l3 3 3-3M12 20v-6M9 17l3-3 3 3"/>', bottom: '<path d="M4 20h16M12 4v12M8 12l4 4 4-4"/>',
  wrap: '<path d="M4 6h16M4 12h12a3 3 0 010 6h-4m0 0l2-2m-2 2l2 2M4 18h4"/>',
  merge: '<rect x="3" y="6" width="18" height="12" rx="1"/><path d="M8 12h8M8 12l2-2M8 12l2 2M16 12l-2-2M16 12l-2 2"/>',
  borders: '<rect x="4" y="4" width="16" height="16"/><path d="M12 4v16M4 12h16" stroke-dasharray="2 2"/>',
  indent: '<path d="M4 5h16M10 10h10M10 15h10M4 20h16M4 9l3 2.5L4 14z"/>', outdent: '<path d="M4 5h16M10 10h10M10 15h10M4 20h16M7 9l-3 2.5L7 14z"/>',
  rows: '<path d="M4 6h16M4 12h16M4 18h16M12 9v6M9 12h6" />',
  band: '<path d="M3 5h18M3 19h18"/><path d="M7 9h10v6H7z" fill="rgba(33,115,70,.15)"/><path d="M12 9V5M12 19v-4"/>',
  img: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M4 18l5-5 4 4 3-3 4 4"/>',
  chart: '<path d="M4 20h16M7 20v-7M12 20V6M17 20v-10"/>',
  cf: '<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M4 10h16M10 4v16"/><path d="M12 12h6v6h-6z" fill="rgba(192,57,43,.35)"/>',
  dv: '<rect x="4" y="5" width="16" height="6" rx="1"/><path d="M8 15l4 4 4-4"/>',
  freeze: '<rect x="4" y="4" width="16" height="16" rx="1"/><path d="M4 9h16M9 4v16" stroke-width="2.2"/>',
  clear: '<path d="M5 19l9-14 5 3-8 12H5z"/><path d="M5 19h14"/>',
  plus: '<path d="M12 5v14M5 12h14"/>', x: '<path d="M6 6l12 12M18 6L6 18"/>', help: '<circle cx="12" cy="12" r="9"/><path d="M9.5 9.5a2.5 2.5 0 115 0c0 1.7-2.5 2-2.5 4M12 17h.01"/>'
};
const svg = (n) => '<svg viewBox="0 0 24 24">' + IC[n] + '</svg>';
UI.svg = svg;

/* ------------------------------------------------------------------ modal / popover helpers */
UI.modal = function (opt) {
  const root = $('#modal-root');
  const ov = el('div', { class: 'overlay' });
  const box = el('div', { class: 'modal' + (opt.wide ? ' wide' : '') });
  const close = () => { ov.remove(); opt.onClose && opt.onClose(); };
  const head = el('div', { class: 'mh' }, opt.title || '', el('button', { class: 'iconbtn x', html: svg('x').replace('<svg', '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"'), onclick: close }));
  const body = el('div', { class: 'mb' }); if (opt.body) body.append(opt.body.nodeType ? opt.body : el('div', { html: opt.body }));
  box.append(head, body);
  if (opt.buttons && opt.buttons.length) {
    const f = el('div', { class: 'mf' });
    for (const b of opt.buttons) f.append(el('button', { class: 'btn' + (b.primary ? ' primary' : '') + (b.danger ? ' danger' : ''), onclick: async () => { if (!b.onclick || (await b.onclick(box)) !== false) close(); } }, b.label));
    box.append(f);
  }
  ov.append(box); ov.addEventListener('mousedown', (e) => { if (e.target === ov && !opt.sticky) close(); });
  root.append(ov);
  const first = box.querySelector('input,select,textarea'); if (first && !opt.noFocus) setTimeout(() => first.focus(), 30);
  return { close, box, body };
};
UI.confirm = (msg, danger) => new Promise(res => { let done = false; UI.modal({ title: msg, body: el('div'), buttons: [{ label: T('ok'), primary: !danger, danger, onclick: () => { done = true; res(true); } }, { label: T('cancel') }], onClose: () => { if (!done) res(false); } }); });
UI.prompt = (title, value, hint) => new Promise(res => {
  const inp = el('input', { type: 'text', value: value || '', style: 'width:100%;padding:7px 10px;border:1px solid var(--line);border-radius:6px' });
  let done = false;
  const m = UI.modal({ title, body: el('div', {}, inp, hint ? el('div', { class: 'help' }, hint) : null), buttons: [{ label: T('ok'), primary: true, onclick: () => { done = true; res(inp.value); } }, { label: T('cancel') }], onClose: () => { if (!done) res(null); } });
  inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { done = true; res(inp.value); m.close(); } });
});

let popAnchor = null;
UI.popover = function (anchor, content, align) {
  const p = $('#popover');
  if (popAnchor === anchor && p.style.display === 'block') { p.style.display = 'none'; popAnchor = null; return; }
  p.innerHTML = ''; p.append(content); p.style.display = 'block'; popAnchor = anchor;
  const r = anchor.getBoundingClientRect(); const pw = p.offsetWidth, ph = p.offsetHeight;
  let x = KXL.RTL ? r.right - pw : r.left; x = Math.max(6, Math.min(x, innerWidth - pw - 6));
  let y = r.bottom + 4; if (y + ph > innerHeight - 6) y = Math.max(6, r.top - ph - 4);
  p.style.left = x + 'px'; p.style.top = y + 'px';
};
UI.closePopover = () => { $('#popover').style.display = 'none'; popAnchor = null; };
document.addEventListener('mousedown', (e) => { const p = $('#popover'); if (p.style.display === 'block' && !p.contains(e.target) && !(popAnchor && popAnchor.contains(e.target))) UI.closePopover(); }, true);

const PALETTE = ['#000000', '#404040', '#7f7f7f', '#bfbfbf', '#e7e6e6', '#ffffff', '#c00000', '#ff0000', '#ffc000', '#ffff00',
  '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0', '#f8cbad', '#fff2cc', '#e2efda', '#ddebf7',
  '#1f4e79', '#2e75b6', '#9dc3e6', '#c55a11', '#f4b183', '#548235', '#a9d18e', '#bf9000', '#ffd966', '#7b7b7b',
  '#d6dce4', '#8497b0', '#44546a', '#e7f1ea', '#217346', '#b4c7e7', '#fce4d6', '#ededed', '#d9d9d9', '#fbe5d6'];
UI.colorPicker = function (anchor, current, cb, allowNone) {
  const box = el('div', {});
  const sw = el('div', { class: 'swatches' });
  PALETTE.forEach(c => sw.append(el('i', { style: 'background:' + c, title: c, onclick: () => { cb(c); UI.closePopover(); } })));
  box.append(sw);
  const row = el('div', { style: 'display:flex;gap:6px;align-items:center' });
  const nat = el('input', { type: 'color', value: /^#[0-9a-f]{6}$/i.test(current || '') ? current : '#217346', style: 'width:40px;height:26px;border:0;padding:0;background:none' });
  nat.addEventListener('change', () => { cb(nat.value); UI.closePopover(); });
  row.append(nat, el('span', { class: 'muted' }, '…'));
  if (allowNone) row.append(el('button', { class: 'btn', style: 'margin-inline-start:auto', onclick: () => { cb(null); UI.closePopover(); } }, '∅'));
  box.append(row);
  UI.popover(anchor, box);
};
UI.menu = function (anchor, items) {
  const box = el('div', { style: 'min-width:190px' });
  for (const it of items) {
    if (it === '-') { box.append(el('div', { class: 'msep' })); continue; }
    box.append(el('button', { class: 'mi' + (it.disabled ? ' dis' : ''), onclick: () => { UI.closePopover(); it.run(); } }, it.icon ? el('span', { html: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.7">' + IC[it.icon] + '</svg>' }) : null, it.label));
  }
  UI.popover(anchor, box);
};

/* ------------------------------------------------------------------ ribbon */
const FONTS = ['Tahoma', 'Segoe UI', 'Arial', 'Calibri', 'Times New Roman', 'Verdana', 'Courier New', 'B Nazanin', 'B Titr', 'B Yekan', 'IRANSans', 'Vazirmatn'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48];
const FORMATS = [['General', 'General'], ['0', '0'], ['0.00', '0.00'], ['#,##0', '#,##0'], ['#,##0.00', '#,##0.00'], ['0%', '0%'], ['0.00%', '0.00%'],
  ['yyyy/mm/dd', 'yyyy/mm/dd'], ['yyyy/mm/dd hh:mm', 'yyyy/mm/dd hh:mm'], ['hh:mm', 'hh:mm'], ['@', '@ (text)']];

function tb(opts) {
  const b = el('button', { class: 'tb ' + (opts.cls || ''), title: opts.title || '', 'data-k': opts.key || '', html: (opts.icon ? svg(opts.icon) : '') + (opts.html || '') + (opts.caret ? '<span class="caret">▾</span>' : '') });
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', (e) => { e.stopPropagation(); opts.run(b, e); });
  return b;
}
function group(...kids) { return el('div', { class: 'rg' }, ...kids); }
const styleNow = () => KXL.styleOf(KXL.cell(S.sel.ar, S.sel.ac));

function buildRibbon() {
  const r = $('#ribbon'); r.innerHTML = '';
  const font = el('input', { class: 'rin', id: 'fontsel', list: 'fontlist', title: T('font') }); font.addEventListener('change', () => G.applyStyle({ font: font.value || undefined }));
  const size = el('input', { class: 'rin', id: 'sizesel', list: 'sizelist', title: T('size'), inputmode: 'decimal' }); size.addEventListener('change', () => { const v = parseFloat(size.value); G.applyStyle({ size: v > 0 ? v : undefined }); });
  const dl1 = el('datalist', { id: 'fontlist' }, ...FONTS.map(f => el('option', { value: f }))), dl2 = el('datalist', { id: 'sizelist' }, ...SIZES.map(f => el('option', { value: f })));
  const toggle = (k) => () => { const on = !styleNow()[k]; G.applyStyle({ [k]: on }); };
  r.append(dl1, dl2,
    group(font, size),
    group(
      tb({ key: 'bold', html: 'B', cls: 'txt', title: T('bold'), run: toggle('bold') }),
      tb({ key: 'italic', html: '<i>I</i>', cls: 'txt', title: T('italic'), run: toggle('italic') }),
      tb({ key: 'underline', html: '<u>U</u>', cls: 'txt', title: T('underline'), run: toggle('underline') }),
      tb({ key: 'strike', html: '<s>S</s>', cls: 'txt', title: T('strike'), run: toggle('strike') }),
      tb({ key: 'color', cls: 'colorbtn', html: '<b>A</b><span class="sw" data-sw="color" style="background:#000"></span>', title: T('fontColor'), caret: false, run: (b) => UI.colorPicker(b, styleNow().color, (c) => G.applyStyle({ color: c || undefined }), true) }),
      tb({ key: 'fill', cls: 'colorbtn', html: '<b style="font-size:12px">▨</b><span class="sw" data-sw="fill" style="background:#ffe699"></span>', title: T('fillColor'), run: (b) => UI.colorPicker(b, styleNow().fill, (c) => G.applyStyle({ fill: c || undefined }), true) }),
      tb({ icon: 'borders', title: T('borders'), caret: true, run: (b) => bordersMenu(b) })
    ),
    group(
      tb({ key: 'h-left', icon: 'left', title: T('left'), run: () => G.applyStyle({ h: styleNow().h === 'left' ? undefined : 'left' }) }),
      tb({ key: 'h-center', icon: 'center', title: T('center'), run: () => G.applyStyle({ h: styleNow().h === 'center' ? undefined : 'center' }) }),
      tb({ key: 'h-right', icon: 'right', title: T('right'), run: () => G.applyStyle({ h: styleNow().h === 'right' ? undefined : 'right' }) }),
      tb({ key: 'v-top', icon: 'top', title: T('top'), run: () => G.applyStyle({ v: styleNow().v === 'top' ? undefined : 'top' }) }),
      tb({ key: 'v-center', icon: 'middle', title: T('middle'), run: () => G.applyStyle({ v: styleNow().v === 'center' ? undefined : 'center' }) }),
      tb({ key: 'v-bottom', icon: 'bottom', title: T('bottom'), run: () => G.applyStyle({ v: styleNow().v === 'bottom' ? undefined : 'bottom' }) }),
      tb({ key: 'wrap', icon: 'wrap', title: T('wrap'), run: toggle('wrap') }),
      tb({ key: 'merge', icon: 'merge', title: T('merge') + ' / ' + T('unmerge'), run: () => { const s = S.sel; const multi = s.r1 !== s.r2 || s.c1 !== s.c2; if (rangeIsMerge() || !multi) G.unmerge(); else G.merge(); } }),
      tb({ icon: 'outdent', title: T('outdent'), run: () => G.applyStyle({ indent: Math.max(0, (styleNow().indent || 0) - 1) }) }),
      tb({ icon: 'indent', title: T('indent'), run: () => G.applyStyle({ indent: Math.min(15, (styleNow().indent || 0) + 1) }) })
    ),
    group((() => { const f = el('select', { class: 'rin', id: 'fmtsel', title: T('numFmt') }); FORMATS.forEach(([v, l]) => f.append(el('option', { value: v }, l))); f.append(el('option', { value: '__custom' }, T('customFmt'))); f.addEventListener('change', async () => { if (f.value === '__custom') { const v = await UI.prompt(T('numFmt'), styleNow().format || '', T('customFmtPrompt')); if (v) G.applyStyle({ format: v }); refreshToolbar(); } else G.applyStyle({ format: f.value === 'General' ? undefined : f.value }); }); return f; })(),
      tb({ html: '<b>٬</b>', title: '#,##0', run: () => G.applyStyle({ format: '#,##0' }) }),
      tb({ html: '<b>%</b>', title: '0%', run: () => G.applyStyle({ format: '0%' }) }),
      tb({ html: '.0←', title: '+', run: () => stepDecimals(1) }), tb({ html: '.0→', title: '−', run: () => stepDecimals(-1) })),
    group(
      tb({ icon: 'rows', title: T('structure'), caret: true, run: (b) => structureMenu(b) }),
      tb({ icon: 'freeze', title: T('freeze'), caret: true, run: (b) => freezeMenu(b) }),
      tb({ icon: 'clear', title: T('clearFmt'), run: () => G.clearFormats() })
    ),
    group(
      tb({ icon: 'band', html: '<span style="font-size:11.5px">' + esc(T('addBand')) + '</span>', title: T('bandHelp'), run: () => bandDialog() }),
      tb({ icon: 'img', title: T('addImage'), run: () => pickImage() }),
      tb({ icon: 'chart', title: T('addChart'), run: () => chartDialog() }),
      tb({ icon: 'cf', title: T('condFmt'), run: () => { S.panelTab = 'rules'; renderRight(); cfDialog(); } }),
      tb({ icon: 'dv', title: T('validation'), run: () => { S.panelTab = 'rules'; renderRight(); dvDialog(); } })
    ));
}
function rangeIsMerge() { const g = KXL.mergeAt(S.sel.ar, S.sel.ac); return g && g.r1 === S.sel.r1 && g.r2 === S.sel.r2 && g.c1 === S.sel.c1 && g.c2 === S.sel.c2; }
function stepDecimals(d) {
  const f = styleNow().format || 'General'; let m = /^(.*?)(0*)(\.(0+))?(.*)$/.exec(f);
  if (f === 'General') { G.applyStyle({ format: d > 0 ? '0.0' : '0' }); return; }
  const dec = m && m[4] ? m[4].length : 0; const n = Math.max(0, dec + d);
  const base = f.replace(/\.0+/, '');
  const at = base.search(/0(?!.*0)/); if (at < 0) return;
  G.applyStyle({ format: base.slice(0, at + 1) + (n ? '.' + '0'.repeat(n) : '') + base.slice(at + 1) });
}
function bordersMenu(anchor) {
  const box = el('div', { style: 'min-width:210px' });
  const mk = (label, mode) => box.append(el('button', { class: 'mi', onclick: () => { UI.closePopover(); G.setBorder(mode); } }, label));
  mk(T('bordersAll'), 'all'); mk(T('bordersOuter'), 'outer'); mk(T('bordersThickOuter'), 'thick'); mk(T('bordersTop'), 'top'); mk(T('bordersBottom'), 'bottom'); mk(T('bordersLeft'), 'left'); mk(T('bordersRight'), 'right'); mk(T('bordersDouble'), 'double'); mk(T('bordersNone'), 'none');
  box.append(el('div', { class: 'msep' }));
  const st = el('select', { class: 'rin', style: 'width:100%;margin-bottom:6px' }, ...['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair'].map(x => el('option', { value: x, selected: S.border.style === x ? 'selected' : null }, x)));
  st.addEventListener('change', () => { S.border.style = st.value; });
  const col = el('input', { type: 'color', value: S.border.color, style: 'width:100%;height:28px' }); col.addEventListener('input', () => { S.border.color = col.value; });
  box.append(el('div', { class: 'lbl2' }, T('bStyle')), st, el('div', { class: 'lbl2' }, T('bColor')), col);
  UI.popover(anchor, box);
}
function structureMenu(anchor) {
  const s = S.sel; const nr = s.r2 - s.r1 + 1, nc = s.c2 - s.c1 + 1;
  UI.menu(anchor, [
    { label: T('insRowAbove'), run: () => insertRows(false) }, { label: T('insRowBelow'), run: () => insertRows(true) }, { label: T('delRow') + ' (' + nr + ')', run: deleteRows }, '-',
    { label: T('insColBefore'), run: () => insertCols(false) }, { label: T('insColAfter'), run: () => insertCols(true) }, { label: T('delCol') + ' (' + nc + ')', run: deleteCols }, '-',
    { label: T('clearCells'), run: () => G.clearContents() }
  ]);
}
function freezeMenu(anchor) {
  const sh = KXL.cur();
  UI.menu(anchor, [
    { label: T('freezeAt'), run: () => { KXL.mutate(() => { sh.freezeRows = S.sel.ar - 1; sh.freezeCols = S.sel.ac - 1; }); G.render(); } },
    { label: T('freeze') + ' ↓ (' + S.sel.ar + ')', run: () => { KXL.mutate(() => { sh.freezeRows = S.sel.r2; }); G.render(); } },
    { label: T('unfreeze'), run: () => { KXL.mutate(() => { sh.freezeRows = 0; sh.freezeCols = 0; }); G.render(); } }
  ]);
}
function insertRows(after) { const s = S.sel, n = s.r2 - s.r1 + 1; KXL.mutate(() => KXL.structural('r', after ? s.r2 + 1 : s.r1, n)); G.render(); }
function deleteRows() { const s = S.sel, n = s.r2 - s.r1 + 1; KXL.mutate(() => KXL.structural('r', s.r1, -1, n)); G.select(s.r1, S.sel.ac); G.render(); }
function insertCols(after) { const s = S.sel, n = s.c2 - s.c1 + 1; KXL.mutate(() => KXL.structural('c', after ? s.c2 + 1 : s.c1, n)); G.render(); }
function deleteCols() { const s = S.sel, n = s.c2 - s.c1 + 1; KXL.mutate(() => KXL.structural('c', s.c1, -1, n)); G.render(); }

function refreshToolbar() {
  const st = styleNow(); const sh = KXL.cur();
  const set = (k, on) => { const b = $('.tb[data-k="' + k + '"]'); if (b) b.classList.toggle('on', !!on); };
  set('bold', st.bold); set('italic', st.italic); set('underline', st.underline); set('strike', st.strike); set('wrap', st.wrap);
  set('h-left', st.h === 'left'); set('h-center', st.h === 'center'); set('h-right', st.h === 'right');
  set('v-top', st.v === 'top'); set('v-center', st.v === 'center'); set('v-bottom', st.v === 'bottom');
  set('merge', rangeIsMerge());
  const f = $('#fontsel'); if (f && document.activeElement !== f) f.value = st.font || S.tpl.defaultFont || 'Tahoma';
  const z = $('#sizesel'); if (z && document.activeElement !== z) z.value = st.size || S.tpl.defaultFontSize || 10;
  const fm = $('#fmtsel'); if (fm) { const v = st.format || 'General'; if (![...fm.options].some(o => o.value === v)) { fm.insertBefore(el('option', { value: v, class: 'tmp' }, v), fm.firstChild); } $$('option.tmp', fm).forEach(o => { if (o.value !== v) o.remove(); }); fm.value = v; }
  const sc = $('[data-sw="color"]'); if (sc) sc.style.background = st.color || '#000'; const sf = $('[data-sw="fill"]'); if (sf) sf.style.background = st.fill || '#ffe699';
}

/* ------------------------------------------------------------------ top bar */
function buildTop() {
  const a = $('#top-actions'); a.innerHTML = '';
  const b = (cls, icon, label, run, title) => el('button', { class: 'btn ' + cls, title: title || '', html: svg(icon) + '<span>' + esc(label) + '</span>', onclick: run });
  a.append(b('', 'new', T('newT'), onNew), b('', 'open', T('open'), onOpen),
    b('', 'undo', '', () => KXL.undoOp(), T('undo')), b('', 'redo', '', () => KXL.redoOp(), T('redo')),
    b('primary', 'save', T('save'), onSave), b('', 'eye', T('preview'), onPreview), b('', 'dl', T('download'), onDownload),
    b('', 'open', 'xlsx', importXlsx, T('importXlsx')),
    S.cfg && S.cfg.pdf ? b('', 'dl', 'PDF', onPdf, T('pdf')) : null,
    b('', 'json', '', exportJson, T('exportJson')), b('', 'json', '', importJson, T('importJson')));
  $$('.btn', a).slice(2, 4).forEach(x => x.querySelector('span').remove());
  $$('.btn', a).slice(-2).forEach(x => x.querySelector('span').remove());
}

/* ------------------------------------------------------------------ formula bar & status */
function updateFormulaBar() {
  const s = S.sel; const c = KXL.cell(s.ar, s.ac);
  $('#cell-addr').textContent = (s.r1 === s.r2 && s.c1 === s.c2) ? KXL.a1(s.ar, s.ac) : KXL.rangeStr(s);
  const fi = $('#formula-input'); if (document.activeElement !== fi && !S.editing) fi.value = c && c.v ? c.v : '';
  const sh = KXL.cur();
  const bands = sh.bands.filter(b => s.ar >= b.firstRow && s.ar <= b.lastRow).map(b => b.name || b.source);
  $('#statusline').textContent = (s.r2 - s.r1 + 1) + '×' + (s.c2 - s.c1 + 1) + (bands.length ? '  ·  ' + T('bandLbl') + ': ' + bands.join(' › ') : '') + (sh.rtl ? '  ·  RTL' : '');
}
function wireFormulaBar() {
  const fi = $('#formula-input');
  const commit = () => { const s = S.sel; const cell = KXL.cell(s.ar, s.ac); const old = cell && cell.v ? cell.v : ''; if (fi.value !== old) { KXL.mutate(() => KXL.setValue(s.ar, s.ac, fi.value)); G.render(); } };
  fi.addEventListener('keydown', (e) => { if (e.key === 'Enter') { commit(); $('#gridwrap').focus(); } else if (e.key === 'Escape') { updateFormulaBar(); $('#gridwrap').focus(); } e.stopPropagation(); });
  fi.addEventListener('change', commit);
  fi.addEventListener('focus', () => fi.select());
  $('#tpl-name').addEventListener('input', (e) => { S.tpl.name = e.target.value; });
}

/* ------------------------------------------------------------------ sheet tabs */
function renderSheetTabs() {
  const host = $('#sheet-tabs'); host.innerHTML = '';
  S.tpl.sheets.forEach((sh, i) => {
    const t = el('div', { class: 'stab' + (i === S.sheetIdx ? ' on' : ''), title: sh.repeatFor ? '↻ ' + sh.repeatFor : '' }, (sh.repeatFor ? '↻ ' : '') + sh.name);
    t.addEventListener('click', () => switchSheet(i));
    t.addEventListener('dblclick', () => renameSheet(i));
    t.addEventListener('contextmenu', (e) => { e.preventDefault(); UI.menu(t, [{ label: T('rename'), run: () => renameSheet(i) }, { label: T('duplicate'), run: () => dupSheet(i) }, { label: T('del'), disabled: S.tpl.sheets.length < 2, run: () => delSheet(i) }]); });
    host.append(t);
  });
  host.append(el('button', { class: 'tb', title: T('addSheet'), html: svg('plus'), onclick: addSheet }));
}
function switchSheet(i) { if (G && S.editing) G.commitEdit(); S.sheetIdx = i; S.sel = { r1: 1, c1: 1, r2: 1, c2: 1, ar: 1, ac: 1 }; renderSheetTabs(); G.render(); renderRight(); updateFormulaBar(); refreshToolbar(); }
function uniqueSheetName(base) { let n = base, i = 1; while (S.tpl.sheets.some(s => s.name === n)) n = base + ' ' + (++i); return n; }
function addSheet() { KXL.mutate(() => { S.tpl.sheets.push(freshSheet(uniqueSheetName(T('sheet')))); }); switchSheet(S.tpl.sheets.length - 1); }
function freshSheet(name) { const sh = { name, rtl: KXL.RTL, showGridLines: true, zoom: 100, cells: [] }; KXL.fn.hydrate(sh); return sh; }
async function renameSheet(i) { const v = await UI.prompt(T('rename'), S.tpl.sheets[i].name, '[]:*?/\\ ✗'); if (v && v.trim()) { KXL.mutate(() => { S.tpl.sheets[i].name = v.trim().slice(0, 31); }); renderSheetTabs(); } }
function dupSheet(i) { KXL.mutate(() => { const d = KXL.serialize().sheets[i]; d.name = uniqueSheetName(d.name); KXL.fn.hydrate(d); S.tpl.sheets.splice(i + 1, 0, d); }); switchSheet(i + 1); }
async function delSheet(i) { if (S.tpl.sheets.length < 2 || !(await UI.confirm(T('confirmDel'), true))) return; KXL.mutate(() => { S.tpl.sheets.splice(i, 1); }); switchSheet(Math.max(0, Math.min(S.sheetIdx, S.tpl.sheets.length - 1))); }

/* ------------------------------------------------------------------ left panel: fields + sample */
let leftTab = 'fields', fieldFilter = '';
function renderLeft() {
  const tabs = $('#left-tabs'); tabs.innerHTML = '';
  [['fields', T('fields')], ['sets', T('sets')], ['sample', T('sample')]].forEach(([k, l]) => tabs.append(el('button', { class: leftTab === k ? 'on' : '', onclick: () => { leftTab = k; renderLeft(); } }, l)));
  const body = $('#left-body'); body.innerHTML = '';
  if (leftTab === 'sample') return renderSample(body);
  if (leftTab === 'sets') return renderSets(body);

  const cfg = S.cfg || { models: [] };
  if (cfg.models.length) {
    const sel = el('select', { class: 'search', title: T('model') }, el('option', { value: '' }, T('noModel')), ...cfg.models.map(m => el('option', { value: m.name, selected: S.tpl.modelName === m.name ? 'selected' : null }, m.name)));
    sel.addEventListener('change', () => { KXL.mutate(() => { S.tpl.modelName = sel.value || undefined; if (sel.value && !S.tpl.sampleData) loadModelSample(sel.value); }); renderLeft(); });
    body.append(sel);
  }
  const search = el('input', { class: 'search', placeholder: T('searchFields'), value: fieldFilter }); search.addEventListener('input', () => { fieldFilter = search.value.toLowerCase(); renderTree(treeHost); }); body.append(search);
  const treeHost = el('div', { class: 'tree' }); body.append(treeHost);
  renderTree(treeHost);
  body.append(el('div', { class: 'hint' }, T('insertHint')));
  const tips = el('div', { class: 'help' }, el('b', {}, T('tipsTitle')), ...T('tips').map(x => el('div', {}, '• ' + x)));
  body.append(tips);
}
async function loadModelSample(name) { try { const m = await KXL.api('GET', 'models/' + encodeURIComponent(name)); KXL.mutate(() => { S.tpl.sampleData = m.sample; }); } catch (e) { /* ignore */ } }

function renderTree(host) {
  host.innerHTML = '';
  const fields = KXL.fieldTree();
  const sp = el('div', { class: 'sec' }, el('div', { class: 'lbl2' }, T('specials')));
  [['$index', '#'], ['$count', '#'], ['$today', 'date'], ['$now', 'date']].forEach(([n, ty]) => sp.append(fnode({ name: n, type: ty === '#' ? 'number' : 'date' }, [], [n], true)));
  host.append(sp);
  const filter = fieldFilter;
  const walk = (nodes, chain, pathArr, into) => {
    for (const f of nodes) {
      const full = pathArr.concat(f.name);
      const kids = f.children || [];
      const match = !filter || f.name.toLowerCase().includes(filter) || hasMatch(f, filter);
      if (!match) continue;
      if (kids.length || f.type === 'list' || f.type === 'object') {
        const node = fnode(f, chain, full, false); const wrapEl = el('div'); wrapEl.append(node);
        const kidsEl = el('div', { class: 'kids' });
        if (filter || node.dataset.open === '1') { kidsEl.style.display = ''; } else kidsEl.style.display = 'none';
        node.dataset.open = filter ? '1' : (node.dataset.open || '0');
        node.querySelector('.tw').textContent = kidsEl.style.display === 'none' ? '▸' : '▾';
        node.addEventListener('click', (e) => { if (e.target.closest('.act')) return; const open = kidsEl.style.display === 'none'; kidsEl.style.display = open ? '' : 'none'; node.querySelector('.tw').textContent = open ? '▾' : '▸'; openState[full.join('.')] = open; });
        if (openState[full.join('.')] && !filter) { kidsEl.style.display = ''; node.querySelector('.tw').textContent = '▾'; }
        walk(kids, chain.concat({ name: f.name, type: f.type }), full, kidsEl);
        wrapEl.append(kidsEl); into.append(wrapEl);
      } else into.append(fnode(f, chain, full, false));
    }
  };
  walk(fields, [], [], host);
  if (!fields.length) host.append(el('div', { class: 'empty' }, T('sampleHelp')));
}
const openState = {};
function hasMatch(f, q) { return (f.children || []).some(c => c.name.toLowerCase().includes(q) || hasMatch(c, q)); }
const TYPE_ICON = { text: 'Aa', number: '#', date: '▦', bool: '✓', list: '☰', object: '{}', image: '▣' };
function fnode(f, chain, full, special) {
  const info = { chain, path: full, leaf: { name: f.name, type: f.type } };
  const isGroup = f.type === 'list' || f.type === 'object';
  const n = el('div', { class: 'fnode', draggable: 'true', title: full.join('.') },
    el('span', { class: 'tw' }, isGroup ? '▸' : ''), el('span', { class: 'ico ' + (special ? 'sp' : f.type) }, TYPE_ICON[f.type] || 'Aa'), el('span', { class: 'nm' }, f.name));
  const act = el('span', { class: 'act' });
  const inList = chain.some(c => c.type === 'list');
  if (f.type === 'number' && inList && !special) act.append(el('button', { title: T('sumBtn'), onclick: (e) => { e.stopPropagation(); insertAtActive(info, 'sum'); } }, 'Σ'));
  if (f.type === 'date') act.append(el('button', { title: T('jalaliBtn'), onclick: (e) => { e.stopPropagation(); insertAtActive(info, 'jalali'); } }, 'ش'));
  n.append(act);
  if (!isGroup) { n.addEventListener('dblclick', () => insertAtActive(info)); }
  else { n.addEventListener('dblclick', (e) => { e.preventDefault(); }); }
  n.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/kxl-field', JSON.stringify(isGroup ? { chain: chain.concat({ name: f.name, type: f.type }), path: full, leaf: { name: f.name, type: 'object' } } : info)); e.dataTransfer.effectAllowed = 'copy'; });
  return n;
}
function insertAtActive(info, mode) {
  const s = S.sel; G.insertField(info, s.ar, s.ac, mode); G.select(s.ar, s.ac); $('#gridwrap').focus();
}

function renderTreeSoon() { /* the field tree is rebuilt when the fields tab opens */ }

function renderSets(body) {
  const t = S.tpl; t.dataSets = t.dataSets || []; t.parameters = t.parameters || [];
  const refresh = () => { renderLeft(); };
  const paths = listPaths(KXL.fieldTree(), '', []);
  body.append(el('datalist', { id: 'srcpaths' }, ...paths.map(p => el('option', { value: p }))));
  body.append(el('div', { class: 'sec' }, el('h4', {}, T('dataSets')), el('button', { class: 'btn', style: 'margin-inline-end:6px', onclick: () => { KXL.mutate(() => { t.dataSets.push({ name: 'Query' + (t.dataSets.length + 1), type: 'sql', connection: (S.cfg.connections || [])[0] || '', sql: 'SELECT ', columns: [] }); }); refresh(); } }, '+ ' + T('addSql')), el('button', { class: 'btn', onclick: () => { KXL.mutate(() => { t.dataSets.push({ name: 'Matrix' + (t.dataSets.length ? t.dataSets.length + 1 : ''), type: 'pivot', source: paths[0] || '', rowPath: '', colPath: '', valuePath: '', agg: 'sum' }); }); refresh(); } }, '+ ' + T('addPivot')),
    el('button', { class: 'btn', style: 'margin-inline-start:6px', onclick: () => { KXL.mutate(() => { t.dataSets.push({ name: 'Data' + (t.dataSets.length + 1), type: 'json', json: '[]', columns: [] }); }); refresh(); } }, '+ ' + T('addJson')),
    el('button', { class: 'btn', style: 'margin-inline-start:6px', onclick: () => { KXL.mutate(() => { t.dataSets.push({ name: 'Data' + (t.dataSets.length + 1), type: 'provider', connection: (S.cfg.providers || [])[0] || '', columns: [] }); }); refresh(); } }, '+ ' + T('addProvider'))));
  body.append(el('div', { class: 'help' }, T('pivotHelp')));
  t.dataSets.forEach((d, i) => {
    const up = (k, v) => KXL.mutate(() => { d[k] = v; });
    if (d.type === 'sql') {
      const sqlBox = el('textarea', { spellcheck: 'false', style: 'direction:ltr;text-align:start;font-family:Consolas,monospace;min-height:90px' }); sqlBox.value = d.sql || '';
      sqlBox.addEventListener('change', () => up('sql', sqlBox.value));
      const msg = el('div', { class: 'muted', style: 'margin-top:6px;direction:ltr;text-align:start;font-size:11.5px' });
      const test = el('button', { class: 'btn', onclick: async () => {
        msg.style.color = ''; msg.textContent = T('loading');
        try {
          const params = {}; (t.parameters || []).forEach(p => { params[p.name] = p.type === 'number' ? Number(p.default) : p.default; });
          const r = await KXL.api('POST', 'sql/test', { connection: d.connection, sql: sqlBox.value, params });
          KXL.mutate(() => { d.sql = sqlBox.value; d.columns = r.columns; });
          msg.textContent = r.count + ' ' + T('rowsFound') + ' · ' + r.columns.map(c => c.name).join(', '); renderTreeSoon();
        } catch (e) { msg.style.color = 'var(--danger)'; msg.textContent = e.message; }
      } }, T('testQuery'));
      body.append(el('div', { class: 'card' }, el('div', { class: 'hd' }, 'SQL · ' + (d.name || ''), el('button', { class: 'iconbtn x', onclick: () => { KXL.mutate(() => { t.dataSets.splice(i, 1); }); refresh(); } }, '\u2715')),
        frow(T('dsName'), inp(d.name, (v) => { up('name', v.trim()); refresh(); }, { dir: 'ltr' })),
        frow(T('dsConnection'), sel((S.cfg.connections || []).map(c => [c, c]).concat((S.cfg.connections || []).includes(d.connection) ? [] : [[d.connection || '', d.connection || '-']]), d.connection || '', (v) => up('connection', v))),
        frow(T('dsSql'), sqlBox), test, msg,
        S.cfg.allowSql ? null : el('div', { class: 'warnbox', style: 'margin-top:8px' }, T('sqlDisabled')), el('div', { class: 'help' }, T('sqlHint'))));
      return;
    }
    if (d.type === 'json' || d.type === 'provider') {
      const isJson = d.type === 'json';
      const msg = el('div', { class: 'muted', style: 'margin-top:6px;direction:ltr;text-align:start;font-size:11.5px' });
      const test = el('button', { class: 'btn', onclick: async () => {
        msg.style.color = ''; msg.textContent = T('loading');
        try {
          const params = {}; (t.parameters || []).forEach(p => { params[p.name] = p.type === 'number' ? Number(p.default) : p.default; });
          const r = await KXL.api('POST', 'data/test', { dataSet: d, params });
          KXL.mutate(() => { d.columns = r.columns; });
          msg.textContent = r.count + ' ' + T('rowsFound') + ' · ' + r.columns.map(c => c.name).join(', ') + (r.warnings && r.warnings.length ? ' ⚠ ' + r.warnings.join('; ') : ''); renderTreeSoon();
        } catch (e) { msg.style.color = 'var(--danger)'; msg.textContent = e.message; }
      } }, T('testData'));
      const names = isJson ? (S.cfg.endpoints || []) : (S.cfg.providers || []);
      const jsonBox = el('textarea', { spellcheck: 'false', style: 'direction:ltr;text-align:start;font-family:Consolas,monospace;min-height:90px' }); jsonBox.value = d.json || '';
      jsonBox.addEventListener('change', () => up('json', jsonBox.value || undefined));
      body.append(el('div', { class: 'card' }, el('div', { class: 'hd' }, (isJson ? 'JSON · ' : 'Provider · ') + (d.name || ''), el('button', { class: 'iconbtn x', onclick: () => { KXL.mutate(() => { t.dataSets.splice(i, 1); }); refresh(); } }, '\u2715')),
        frow(T('dsName'), inp(d.name, (v) => { up('name', v.trim()); refresh(); }, { dir: 'ltr' })),
        isJson ? frow(T('dsJsonInline'), jsonBox) : null,
        frow(isJson ? T('dsEndpoint') : T('dsProvider'), sel([['', '—']].concat(names.map(c => [c, c])).concat(!d.connection || names.includes(d.connection) ? [] : [[d.connection, d.connection]]), d.connection || '', (v) => up('connection', v || undefined))),
        isJson ? el('div', { class: 'two' }, frow(T('dsUrl'), inp(d.url || '', (v) => up('url', v.trim() || undefined), { dir: 'ltr', placeholder: '/orders?year={Year}' })), frow(T('dsJsonPath'), inp(d.jsonPath || '', (v) => up('jsonPath', v.trim() || undefined), { dir: 'ltr', placeholder: 'data.items' }))) : null,
        test, msg,
        el('div', { class: 'help' }, isJson ? T('jsonHelp') : T('providerHelp')),
        (names.length || (isJson && d.json)) ? null : el('div', { class: 'warnbox', style: 'margin-top:8px' }, isJson ? T('noEndpoints') : T('noProviders'))));
      return;
    }
    body.append(el('div', { class: 'card' }, el('div', { class: 'hd' }, d.name || 'pivot', el('button', { class: 'iconbtn x', onclick: () => { KXL.mutate(() => { t.dataSets.splice(i, 1); }); refresh(); } }, '\u2715')),
      frow(T('dsName'), inp(d.name, (v) => { up('name', v.trim()); refresh(); }, { dir: 'ltr' })),
      frow(T('dsSource'), inp(d.source, (v) => { up('source', v.trim()); }, { dir: 'ltr', list: 'srcpaths' })),
      el('div', { class: 'two' }, frow(T('dsRow'), inp(d.rowPath, (v) => up('rowPath', v.trim()), { dir: 'ltr' })), frow(T('dsCol'), inp(d.colPath, (v) => up('colPath', v.trim()), { dir: 'ltr' }))),
      el('div', { class: 'two' }, frow(T('dsValue'), inp(d.valuePath, (v) => up('valuePath', v.trim()), { dir: 'ltr' })), frow(T('dsAgg'), sel(['sum', 'count', 'avg', 'min', 'max'].map(x => [x, x]), d.agg || 'sum', (v) => up('agg', v))))));
  });
  body.append(el('div', { class: 'sec', style: 'margin-top:14px' }, el('h4', {}, T('params')), el('button', { class: 'btn', onclick: () => { KXL.mutate(() => { t.parameters.push({ name: 'Param' + (t.parameters.length + 1), type: 'text', label: '', default: '' }); }); refresh(); } }, '+ ' + T('addParam'))));
  body.append(el('div', { class: 'help' }, T('paramHelp')));
  t.parameters.forEach((p, i) => {
    const up = (k, v) => KXL.mutate(() => { p[k] = v; });
    body.append(el('div', { class: 'card' }, el('div', { class: 'hd' }, '{{$p.' + p.name + '}}', el('button', { class: 'iconbtn x', onclick: () => { KXL.mutate(() => { t.parameters.splice(i, 1); }); refresh(); } }, '\u2715')),
      el('div', { class: 'two' }, frow(T('pName'), inp(p.name, (v) => { up('name', v.trim()); refresh(); }, { dir: 'ltr' })), frow(T('pType'), sel([['text', 'text'], ['number', 'number'], ['date', 'date'], ['bool', 'bool']], p.type || 'text', (v) => up('type', v)))),
      el('div', { class: 'two' }, frow(T('pLabel'), inp(p.label || '', (v) => up('label', v))), frow(T('pDefault'), inp(p.default || '', (v) => up('default', v), { dir: 'ltr' })))));
  });
}

function renderSample(body) {
  body.append(el('div', { class: 'help' }, T('sampleHelp')));
  const ta = el('textarea', { style: 'width:100%;min-height:300px;font-family:Consolas,monospace;font-size:12px;direction:ltr;padding:8px;border:1px solid var(--line);border-radius:6px', spellcheck: 'false' });
  ta.value = S.tpl.sampleData ? JSON.stringify(S.tpl.sampleData, null, 2) : '';
  const msg = el('div', { class: 'muted', style: 'margin:6px 0' });
  body.append(ta, msg, el('button', { class: 'btn primary', onclick: () => {
    try { const v = ta.value.trim() ? JSON.parse(ta.value) : undefined; KXL.mutate(() => { S.tpl.sampleData = v; if (v && !S.cfg.models.some(m => m.name === S.tpl.modelName)) S.tpl.modelName = undefined; }); leftTab = 'fields'; renderLeft(); KXL.toast(T('saved'), 'ok'); }
    catch (e) { msg.textContent = T('invalidJson') + ': ' + e.message; msg.style.color = 'var(--danger)'; }
  } }, T('applyJson')));
}

/* ------------------------------------------------------------------ right panel */
const RTABS = [['cell', 'cell'], ['sheet', 'sheetP'], ['bands', 'bands'], ['rules', 'rules'], ['page', 'page'], ['images', 'images']];
function renderRight() {
  const tabs = $('#right-tabs'); tabs.innerHTML = '';
  RTABS.forEach(([k, l]) => tabs.append(el('button', { class: S.panelTab === k ? 'on' : '', onclick: () => { S.panelTab = k; renderRight(); } }, T(l))));
  const body = $('#right-body'); body.innerHTML = '';
  ({ cell: tabCell, sheet: tabSheet, bands: tabBands, rules: tabRules, page: tabPage, images: tabImages }[S.panelTab])(body);
}
const frow = (label, input, extra) => el('div', { class: 'frow' }, el('label', {}, label), input, extra || null);
const inp = (value, onchange, attrs) => { const i = el('input', Object.assign({ type: 'text', value: value == null ? '' : value }, attrs || {})); i.addEventListener('change', () => onchange(i.value)); return i; };
const chk = (label, val, onchange) => { const i = el('input', { type: 'checkbox' }); i.checked = !!val; i.addEventListener('change', () => onchange(i.checked)); return el('label', { class: 'chk' }, i, label); };
const sel = (options, val, onchange) => { const s = el('select', {}, ...options.map(([v, l]) => el('option', { value: v, selected: String(v) === String(val) ? 'selected' : null }, l))); s.addEventListener('change', () => onchange(s.value)); return s; };

function tabCell(body) {
  const s = S.sel; const c = KXL.cell(s.ar, s.ac) || {};
  body.append(el('div', { class: 'sec' }, el('h4', {}, T('cell') + ' ' + KXL.a1(s.ar, s.ac))));
  const ta = el('textarea', { spellcheck: 'false', style: 'direction:ltr;text-align:start;font-family:Consolas,monospace' }); ta.value = c.v || '';
  ta.addEventListener('change', () => { KXL.mutate(() => KXL.setValue(s.ar, s.ac, ta.value)); G.render(); updateFormulaBar(); });
  body.append(frow(T('value'), ta));
  body.append(frow(T('valueType'), sel([['', T('typeAuto')], ['text', T('typeText')]], c.t || '', (v) => { KXL.mutate(() => { const x = KXL.ensureCell(s.ar, s.ac); if (v) x.t = v; else delete x.t; }); })));
  body.append(frow(T('link'), inp(c.l || '', (v) => { KXL.mutate(() => { const x = KXL.ensureCell(s.ar, s.ac); if (v) x.l = v; else delete x.l; }); G.render(); }, { dir: 'ltr', placeholder: 'https://… | #Sheet2!A1' })));
  body.append(frow(T('barcode'), sel([['', T('bcNone')], ['qr', T('bcQr')], ['qr:L', T('bcQr') + ' (L)'], ['qr:H', T('bcQr') + ' (H)'], ['code128', T('bcCode128')]], c.bc || '', (v) => { KXL.mutate(() => { const x = KXL.ensureCell(s.ar, s.ac); if (v) x.bc = v; else delete x.bc; }); G.render(); renderRight(); })));
  if (c.bc) body.append(el('div', { class: 'help' }, T('bcHelp')));
  body.append(chk(T('rtlText'), styleNow().rtl === true, (v) => G.applyStyle({ rtl: v ? true : undefined })));
  body.append(el('div', { class: 'help' }, T('hintFormula')));
  body.append(el('div', { class: 'help', dir: 'ltr', style: 'text-align:left' }, T('exprHint')));
  const filters = el('div', { class: 'help', dir: 'ltr', style: 'text-align:left' }, el('b', {}, 'Filters: '), '{{Date|fa}} · {{Price|round:2}} · {{Name|upper}} · {{X|default:-}} · {{Flag|yesno:Yes/No}}', el('br'), el('b', {}, 'Aggregates: '), '{{Count(Items)}} · {{Sum(Items.Price)}}');
  body.append(filters);
}

function tabSheet(body) {
  const sh = KXL.cur(); const ch = () => { renderSheetTabs(); };
  body.append(frow(T('sheetName'), inp(sh.name, (v) => { if (v.trim()) KXL.mutate(() => { sh.name = v.trim().slice(0, 31); }); ch(); })));
  body.append(chk(T('rtl'), sh.rtl, (v) => { KXL.mutate(() => { sh.rtl = v; }); G.render(); }));
  body.append(chk(T('gridlines'), sh.showGridLines, (v) => { KXL.mutate(() => { sh.showGridLines = v; }); G.render(); }));
  body.append(chk(T('protect'), sh.protect, (v) => KXL.mutate(() => { sh.protect = v; })));
  const col = el('input', { type: 'color', value: sh.tabColor || '#217346', style: 'width:100%;height:30px' }); col.addEventListener('change', () => KXL.mutate(() => { sh.tabColor = col.value; }));
  body.append(frow(T('tabColor'), col));
  body.append(el('div', { class: 'two' }, frow(T('zoom') + ' %', inp(sh.zoom, (v) => KXL.mutate(() => { sh.zoom = Math.max(10, Math.min(400, parseInt(v) || 100)); }), { type: 'number' })), frow(T('defCol'), inp(sh.defaultColWidth, (v) => { KXL.mutate(() => { sh.defaultColWidth = Math.max(2, parseFloat(v) || 12); }); G.render(); }, { type: 'number' }))));
  body.append(frow(T('defRow'), inp(sh.defaultRowHeight, (v) => { KXL.mutate(() => { sh.defaultRowHeight = Math.max(8, parseFloat(v) || 18); }); G.render(); }, { type: 'number' })));
  body.append(frow(T('repeatFor'), inp(sh.repeatFor || '', (v) => { KXL.mutate(() => { sh.repeatFor = v.trim() || undefined; }); ch(); }, { dir: 'ltr', placeholder: 'Branches' })));
  const f = el('div', { class: 'frow' }, el('label', {}, T('autoFilter') + (sh.autoFilter ? ' — ' + sh.autoFilter : '')),
    el('div', { style: 'display:flex;gap:6px' }, el('button', { class: 'btn', onclick: () => { KXL.mutate(() => { sh.autoFilter = KXL.rangeStr({ r1: S.sel.r1, c1: S.sel.c1, r2: S.sel.r1, c2: S.sel.c2 }); }); renderRight(); } }, T('add')), sh.autoFilter ? el('button', { class: 'btn', onclick: () => { KXL.mutate(() => { sh.autoFilter = undefined; }); renderRight(); } }, T('clearFilter')) : null));
  body.append(f);
  body.append(el('div', { class: 'two' }, frow('Freeze rows', inp(sh.freezeRows || 0, (v) => { KXL.mutate(() => { sh.freezeRows = Math.max(0, parseInt(v) || 0); }); G.render(); }, { type: 'number' })), frow('Freeze cols', inp(sh.freezeCols || 0, (v) => { KXL.mutate(() => { sh.freezeCols = Math.max(0, parseInt(v) || 0); }); G.render(); }, { type: 'number' }))));
  body.append(el('hr', { style: 'border:0;border-top:1px solid var(--line2);margin:12px 0' }));
  body.append(frow('Default font', inp(S.tpl.defaultFont || 'Tahoma', (v) => KXL.mutate(() => { S.tpl.defaultFont = v || 'Tahoma'; })), null));
  body.append(frow('Default size', inp(S.tpl.defaultFontSize || 10, (v) => KXL.mutate(() => { S.tpl.defaultFontSize = parseFloat(v) || 10; }), { type: 'number' })));
  body.append(frow(T('templateName') + ' / title', inp(S.tpl.title || '', (v) => KXL.mutate(() => { S.tpl.title = v; }))));
}

function listPaths(nodes, prefix, out) { for (const f of nodes || []) { const p = prefix ? prefix + '.' + f.name : f.name; if (f.type === 'list') out.push(p); if (f.children) listPaths(f.children, p, out); } return out; }

function tabBands(body) {
  const sh = KXL.cur();
  body.append(el('div', { class: 'sec' }, el('h4', {}, T('sections')),
    el('div', { style: 'display:flex;gap:6px;flex-wrap:wrap' },
      ...[['reportHeader', 'secReportHeader'], ['pageHeader', 'secPageHeader'], ['reportFooter', 'secReportFooter']].map(([k, l]) =>
        el('button', { class: 'btn', title: T('addSection'), onclick: () => { KXL.mutate(() => { sh.sections = sh.sections.filter(x => x.kind !== k || k === 'reportFooter'); sh.sections.push({ kind: k, name: T(l), firstRow: S.sel.r1, lastRow: S.sel.r2 }); }); G.render(); renderRight(); } }, '+ ' + T(l)))),
    el('div', { class: 'help' }, T('secHelp'))));
  (sh.sections || []).forEach((x, i) => body.append(el('div', { class: 'card' }, el('div', { class: 'hd' }, T('sec' + x.kind[0].toUpperCase() + x.kind.slice(1)), el('span', { class: 'pill' }, x.firstRow + '–' + x.lastRow),
    el('button', { class: 'iconbtn x', onclick: () => { KXL.mutate(() => { sh.sections.splice(i, 1); }); G.render(); renderRight(); } }, '\u2715')))));
  body.append(el('hr', { style: 'border:0;border-top:1px solid var(--line2);margin:12px 0' }));
  body.append(el('button', { class: 'btn primary', style: 'margin-bottom:8px;margin-inline-end:6px', onclick: () => groupWizard() }, '\u2728 ' + T('groupWizard')));
  body.append(el('div', { class: 'help' }, T('bandHelp')));
  body.append(el('button', { class: 'btn primary', style: 'margin-bottom:10px', onclick: () => bandDialog() }, '+ ' + T('addBand')));
  const paths = listPaths(KXL.fieldTree(), '', []);
  body.append(el('datalist', { id: 'listpaths' }, ...paths.map(p => el('option', { value: p }))));
  if (!sh.bands.length) { body.append(el('div', { class: 'empty' }, T('noBands'))); return; }
  sh.bands.forEach((b, i) => {
    const isCol = b.orientation === 'columns';
    const card = el('div', { class: 'card' }, el('div', { class: 'hd' }, el('span', { class: 'pill', style: 'background:' + ['#2e75b6', '#c55a11', '#548235', '#7030a0', '#bf9000', '#c00000'][i % 6] + ';color:#fff' }, b.name || b.source),
      el('button', { class: 'iconbtn x', title: T('del'), onclick: () => { KXL.mutate(() => { sh.bands.splice(i, 1); }); G.render(); renderRight(); } }, '✕')));
    card.append(frow(T('bandName'), inp(b.name, (v) => { KXL.mutate(() => { b.name = v.trim(); }); G.render(); }, { dir: 'ltr' })));
    card.append(frow(T('orientation2'), sel([['', T('orRows')], ['columns', T('orCols')]], b.orientation || '', (v) => { KXL.mutate(() => { b.orientation = v || undefined; if (v === 'columns' && !b.firstCol) { b.firstCol = S.sel.c1; b.lastCol = S.sel.c2; } }); G.render(); renderRight(); })));
    if (isCol) card.append(el('div', { class: 'help' }, T('colBandHelp')));
    card.append(frow(T('bandSource'), inp(b.source, (v) => { KXL.mutate(() => { b.source = v.trim(); }); G.render(); }, { dir: 'ltr', list: 'listpaths' })));
    card.append(el('div', { class: 'two' }, frow(T('sortBy'), inp(b.sortBy || '', (v) => KXL.mutate(() => { b.sortBy = v.trim() || undefined; }), { dir: 'ltr', placeholder: 'Date' })), frow(T('groupBy'), inp(b.groupBy || '', (v) => KXL.mutate(() => { b.groupBy = v.trim() || undefined; }), { dir: 'ltr', placeholder: 'Category' }))));
    card.append(chk(T('sortDesc'), b.sortDesc, (v) => KXL.mutate(() => { b.sortDesc = v || undefined; })));
    if (b.groupBy) card.append(el('div', { class: 'help' }, T('groupHelp')));
    if (isCol) card.append(el('div', { class: 'two' }, frow(T('from') + ' (' + T('colsRange') + ')', inp(b.firstCol || 1, (v) => { KXL.mutate(() => { b.firstCol = Math.max(1, parseInt(v) || 1); if ((b.lastCol || 0) < b.firstCol) b.lastCol = b.firstCol; }); G.render(); renderRight(); }, { type: 'number' })), frow(T('to'), inp(b.lastCol || 1, (v) => { KXL.mutate(() => { b.lastCol = Math.max(b.firstCol || 1, parseInt(v) || 1); }); G.render(); renderRight(); }, { type: 'number' }))));
    else card.append(el('div', { class: 'two' }, frow(T('from'), inp(b.firstRow, (v) => { KXL.mutate(() => { b.firstRow = Math.max(1, parseInt(v) || 1); if (b.lastRow < b.firstRow) b.lastRow = b.firstRow; }); G.render(); renderRight(); }, { type: 'number' })), frow(T('to'), inp(b.lastRow, (v) => { KXL.mutate(() => { b.lastRow = Math.max(b.firstRow, parseInt(v) || b.firstRow); }); G.render(); renderRight(); }, { type: 'number' }))));
    const alt = el('input', { type: 'color', value: b.alternateFill || '#f2f7fc', style: 'width:42px;height:26px' }); alt.addEventListener('change', () => KXL.mutate(() => { b.alternateFill = alt.value; }));
    card.append(el('div', { class: 'frow' }, el('label', {}, T('bandAlt')), el('div', { style: 'display:flex;gap:6px;align-items:center' }, alt, el('button', { class: 'btn', onclick: () => { KXL.mutate(() => { b.alternateFill = undefined; }); renderRight(); } }, '∅'), b.alternateFill ? el('span', { class: 'pill' }, b.alternateFill) : null)));
    card.append(chk(T('bandKeep'), b.keepWhenEmpty, (v) => KXL.mutate(() => { b.keepWhenEmpty = v; })));
    card.append(chk(T('pageBreak'), b.pageBreak, (v) => KXL.mutate(() => { b.pageBreak = v || undefined; })));
    body.append(card);
  });
}

/* ---- group wizard: builds a full grouped table (titles + group header + detail + group total) */
function findNode(nodes, path) {
  let cur = { children: nodes };
  for (const seg of path.split('.')) { cur = (cur.children || []).find(x => x.name === seg); if (!cur) return null; }
  return cur;
}
function groupWizard() {
  const tree = KXL.fieldTree(); const paths = listPaths(tree, '', []);
  if (!paths.length) { KXL.toast(T('gwNeedList'), 'err'); return; }
  const listSel = sel(paths.map(p => [p, p]), paths[0], () => fill());
  const groupSel = el('select', { class: 'rin', style: 'width:100%' });
  const fieldsBox = el('div', { style: 'display:grid;grid-template-columns:1fr 1fr;gap:4px 12px' });
  const checks = [];
  function fill() {
    const node = findNode(tree, listSel.value); const leaves = ((node && node.children) || []).filter(f => !f.children || !f.children.length);
    groupSel.innerHTML = ''; leaves.filter(f => f.type !== 'number').concat(leaves.filter(f => f.type === 'number')).forEach(f => groupSel.append(el('option', { value: f.name }, f.name)));
    fieldsBox.innerHTML = ''; checks.length = 0;
    leaves.slice(0, 14).forEach((f, i) => {
      const c = el('input', { type: 'checkbox' }); c.checked = i < 6;
      const t = el('input', { type: 'checkbox', title: T('gwTotals') }); t.checked = f.type === 'number';
      checks.push({ f, c, t });
      fieldsBox.append(el('label', { class: 'chk', style: 'margin:0' }, c, f.name + ' ', el('span', { class: 'muted' }, f.type === 'number' ? '(' : ''), f.type === 'number' ? t : null, f.type === 'number' ? el('span', { class: 'muted' }, '\u03A3)') : null));
    });
  }
  fill();
  const body = el('div', {}, el('div', { class: 'help' }, T('gwHelp')), frow(T('gwList'), listSel), frow(T('gwGroup'), groupSel), el('div', { class: 'lbl2' }, T('gwFields') + ' (\u03A3 = ' + T('gwTotals') + ')'), fieldsBox);
  UI.modal({ title: T('groupWizard'), body, buttons: [{ label: T('add'), primary: true, onclick: () => {
    const picked = checks.filter(x => x.c.checked); if (!picked.length) { KXL.toast(T('gwFields'), 'err'); return false; }
    const sh = KXL.cur(); const r = S.sel.r1, c0 = S.sel.c1; const n = picked.length;
    const border = (col) => ({ style: 'thin', color: col });
    const sHead = { bold: true, color: '#FFFFFF', fill: '#2E75B6', h: 'center', v: 'center', left: border('#1F4E79'), right: border('#1F4E79'), top: border('#1F4E79'), bottom: border('#1F4E79') };
    const sGroup = { bold: true, color: '#1F4E79', fill: '#DDEBF7', v: 'center', bottom: border('#9DC3E6') };
    const sCell = { v: 'center', bottom: border('#D0D7DE') };
    const sNum = Object.assign({}, sCell, { format: '#,##0' });
    const sTot = { bold: true, fill: '#F2F2F2', format: '#,##0', top: border('#7F7F7F'), v: 'center' };
    KXL.mutate(() => {
      KXL.structural('r', r, 4);
      const put = (row, col, v, st) => { KXL.setValue(row, col, v); const x = KXL.ensureCell(row, col); x.s = KXL.styleId(st); };
      picked.forEach((p, i) => {
        const col = c0 + i, isNum = p.f.type === 'number';
        put(r, col, p.f.name, sHead);
        put(r + 1, col, i === 0 ? '{{Key}} ({{Count}})' : '', sGroup);
        put(r + 2, col, '{{' + p.f.name + '}}', isNum ? sNum : sCell);
        put(r + 3, col, i === 0 ? T('gwTotalLbl') : (isNum && p.t.checked ? '=SUM([Detail.' + p.f.name + '])' : ''), i === 0 ? Object.assign({}, sTot, { format: undefined }) : sTot);
        if (!sh._colw.get(col)) sh._colw.set(col, isNum ? 14 : 22);
      });
      if (n > 1) sh.merges.push(KXL.rangeStr({ r1: r + 1, c1: c0, r2: r + 1, c2: c0 + n - 1 }));
      const gname = G.uniqueBandName(sh, 'Groups'), dname = G.uniqueBandName(sh, 'Detail');
      sh.bands.push({ name: gname, source: listSel.value, groupBy: groupSel.value, sortBy: groupSel.value, firstRow: r + 1, lastRow: r + 3 });
      sh.bands.push({ name: dname, source: 'Items', firstRow: r + 2, lastRow: r + 2, alternateFill: '#F7FAFD' });
      if (dname !== 'Detail') picked.forEach((p, i) => { const x = KXL.cell(r + 3, c0 + i); if (x && x.v && x.v.includes('[Detail.')) x.v = x.v.replace('[Detail.', '[' + dname + '.'); });
    });
    G.render(); S.panelTab = 'bands'; renderRight(); G.select(r, c0);
  } }, { label: T('cancel') }] });
}

function bandDialog() {
  const sh = KXL.cur(); const s = S.sel;
  const paths = listPaths(KXL.fieldTree(), '', []);
  const name = el('input', { type: 'text', dir: 'ltr' }), src = el('input', { type: 'text', dir: 'ltr', list: 'listpaths2' });
  const r1 = el('input', { type: 'number', value: s.r1 }), r2 = el('input', { type: 'number', value: s.r2 });
  const dl = el('datalist', { id: 'listpaths2' }, ...paths.map(p => el('option', { value: p })));
  src.addEventListener('input', () => { if (!name.dataset.touched) name.value = G.uniqueBandName(sh, (src.value.split('.').pop() || 'Band')); }); name.addEventListener('input', () => { name.dataset.touched = '1'; });
  if (paths.length) { src.value = paths[0]; name.value = G.uniqueBandName(sh, paths[0].split('.').pop()); }
  const alt = el('input', { type: 'color', value: '#f2f7fc' }); const useAlt = el('input', { type: 'checkbox' }); useAlt.checked = true;
  const body = el('div', {}, dl, el('div', { class: 'help' }, T('bandHelp')),
    frow(T('bandSource'), src), frow(T('bandName'), name), el('div', { class: 'two' }, frow(T('from'), r1), frow(T('to'), r2)),
    el('label', { class: 'chk' }, useAlt, T('bandAlt'), alt));
  UI.modal({ title: T('addBand'), body, buttons: [{ label: T('add'), primary: true, onclick: () => {
    const source = src.value.trim(); if (!source) { KXL.toast(T('bandSource'), 'err'); return false; }
    const a = Math.max(1, parseInt(r1.value) || 1), b = Math.max(a, parseInt(r2.value) || a);
    KXL.mutate(() => { sh.bands.push({ name: (name.value || G.uniqueBandName(sh, source.split('.').pop())).trim(), source, firstRow: a, lastRow: b, alternateFill: useAlt.checked ? alt.value : undefined }); });
    G.render(); S.panelTab = 'bands'; renderRight();
  } }, { label: T('cancel') }] });
}

/* ---- rules (conditional formats + validations) */
function opLabel(o) { return { greaterThan: T('opGt'), lessThan: T('opLt'), greaterThanOrEqual: T('opGe'), lessThanOrEqual: T('opLe'), equal: T('opEq'), notEqual: T('opNe'), between: T('opBetween'), notBetween: T('opNotBetween') }[o] || o; }
const OPS = ['greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'equal', 'notEqual', 'between', 'notBetween'];
function tabRules(body) {
  const sh = KXL.cur();
  body.append(el('div', { class: 'sec' }, el('h4', {}, T('condFmt')), el('button', { class: 'btn', onclick: () => cfDialog() }, '+ ' + T('add'))));
  if (!sh.conditionalFormats.length) body.append(el('div', { class: 'muted', style: 'margin-bottom:12px' }, T('noRules')));
  sh.conditionalFormats.forEach((cf, i) => {
    const sw = cf.style && cf.style.fill ? '<i style="display:inline-block;width:12px;height:12px;border-radius:3px;background:' + cf.style.fill + ';margin-inline-end:4px"></i>' : '';
    body.append(el('div', { class: 'card' }, el('div', { class: 'hd', html: sw + '<span dir="ltr">' + esc(cf.range) + '</span> <span class="pill">' + esc(cf.type) + '</span>' }, el('button', { class: 'iconbtn x', title: T('edit'), onclick: () => cfDialog(i) }, '✎'), el('button', { class: 'iconbtn', title: T('del'), onclick: () => { KXL.mutate(() => { sh.conditionalFormats.splice(i, 1); }); renderRight(); } }, '✕')),
      el('div', { class: 'muted', dir: 'ltr', style: 'text-align:start' }, [cf.operator && opLabel(cf.operator), cf.formula1, cf.formula2, cf.text].filter(Boolean).join(' ') )));
  });
  body.append(el('div', { class: 'sec', style: 'margin-top:14px' }, el('h4', {}, T('validation')), el('button', { class: 'btn', onclick: () => dvDialog() }, '+ ' + T('add'))));
  if (!sh.validations.length) body.append(el('div', { class: 'muted' }, T('noRules')));
  sh.validations.forEach((v, i) => body.append(el('div', { class: 'card' }, el('div', { class: 'hd', html: '<span dir="ltr">' + esc(v.range) + '</span> <span class="pill">' + esc(v.type) + '</span>' }, el('button', { class: 'iconbtn x', title: T('edit'), onclick: () => dvDialog(i) }, '✎'), el('button', { class: 'iconbtn', onclick: () => { KXL.mutate(() => { sh.validations.splice(i, 1); }); renderRight(); } }, '✕')), el('div', { class: 'muted', dir: 'ltr', style: 'text-align:start' }, v.formula1 || ''))));
}

function colorInput(val, fallback) { return el('input', { type: 'color', value: val || fallback, style: 'width:44px;height:28px' }); }
function cfDialog(idx) {
  const sh = KXL.cur(); const old = idx != null ? sh.conditionalFormats[idx] : null;
  const range = el('input', { type: 'text', dir: 'ltr', value: old ? old.range : KXL.rangeStr(S.sel) });
  const type = sel([['cellIs', T('cfCellIs')], ['expression', T('cfExpr')], ['containsText', T('cfText')], ['duplicates', T('cfDup')], ['top10', T('cfTop')], ['colorScale', T('cfScale')], ['dataBar', T('cfBar')]], old ? old.type : 'cellIs', () => sync());
  const op = sel(OPS.map(o => [o, opLabel(o)]), old ? old.operator || 'greaterThan' : 'greaterThan', () => sync());
  const v1 = el('input', { type: 'text', dir: 'ltr', value: old ? old.formula1 || '' : '' }), v2 = el('input', { type: 'text', dir: 'ltr', value: old ? old.formula2 || '' : '' });
  const text = el('input', { type: 'text', value: old ? old.text || '' : '' });
  const rank = el('input', { type: 'number', value: old ? old.rank || 10 : 10 }); const bottom = el('input', { type: 'checkbox' }); bottom.checked = !!(old && old.bottom);
  const st = (old && old.style) || { fill: '#ffc7ce', color: '#9c0006' };
  const fill = colorInput(st.fill, '#ffc7ce'), fcol = colorInput(st.color, '#9c0006'), bold = el('input', { type: 'checkbox' }); bold.checked = !!st.bold;
  const useFill = el('input', { type: 'checkbox' }); useFill.checked = st.fill !== undefined; const useCol = el('input', { type: 'checkbox' }); useCol.checked = st.color !== undefined;
  const c1 = colorInput(old && old.colors && old.colors[0], '#f8696b'), c2 = colorInput(old && old.colors && old.colors[1], '#ffeb84'), c3 = colorInput(old && old.colors && old.colors[2], '#63be7b'); const bar = colorInput(old && old.barColor, '#638ec6');
  const rOp = frow(T('operator'), op), rV1 = frow(T('value1'), v1), rV2 = frow(T('value2'), v2), rText = frow(T('textLbl'), text), rRank = frow('N', rank), rBottom = el('label', { class: 'chk' }, bottom, 'Bottom');
  const rFmt = el('div', { class: 'frow' }, el('label', {}, T('format')), el('div', { style: 'display:flex;gap:10px;align-items:center;flex-wrap:wrap' }, el('label', { class: 'chk' }, useFill, T('fill'), fill), el('label', { class: 'chk' }, useCol, T('fontColor'), fcol), el('label', { class: 'chk' }, bold, T('bold').replace(/ \(.*/, ''))));
  const rScale = el('div', { class: 'frow' }, el('label', {}, T('cfScale')), el('div', { style: 'display:flex;gap:8px' }, c1, c2, c3));
  const rBar = frow(T('cfBar'), bar);
  const f = (v, show) => { v.style.display = show ? '' : 'none'; };
  function sync() {
    const t = type.value, between = op.value === 'between' || op.value === 'notBetween';
    f(rOp, t === 'cellIs'); f(rV1, t === 'cellIs' || t === 'expression'); f(rV2, t === 'cellIs' && between); f(rText, t === 'containsText'); f(rRank, t === 'top10'); f(rBottom, t === 'top10');
    f(rFmt, !['colorScale', 'dataBar'].includes(t)); f(rScale, t === 'colorScale'); f(rBar, t === 'dataBar');
    v1.placeholder = t === 'expression' ? 'A1>100' : '100';
  }
  const body = el('div', {}, frow(T('rangeLbl'), range), frow(T('type'), type), rOp, rV1, rV2, rText, rRank, rBottom, rFmt, rScale, rBar, el('div', { class: 'help' }, 'Relative references (e.g. A1>100) refer to the first cell of the range.'));
  sync();
  UI.modal({ title: T('condFmt'), body, buttons: [{ label: old ? T('ok') : T('add'), primary: true, onclick: () => {
    if (!KXL.parseRange(range.value)) { KXL.toast(T('rangeLbl'), 'err'); return false; }
    const t = type.value; const cf = { range: range.value.trim().toUpperCase(), type: t };
    if (t === 'cellIs') { cf.operator = op.value; cf.formula1 = v1.value; if (op.value === 'between' || op.value === 'notBetween') cf.formula2 = v2.value; }
    if (t === 'expression') cf.formula1 = v1.value; if (t === 'containsText') cf.text = text.value; if (t === 'top10') { cf.rank = parseInt(rank.value) || 10; cf.bottom = bottom.checked; }
    if (t === 'colorScale') cf.colors = [c1.value, c2.value, c3.value]; if (t === 'dataBar') cf.barColor = bar.value;
    if (!['colorScale', 'dataBar'].includes(t)) { const s = {}; if (useFill.checked) s.fill = fill.value; if (useCol.checked) s.color = fcol.value; if (bold.checked) s.bold = true; cf.style = s; }
    KXL.mutate(() => { if (old) sh.conditionalFormats[idx] = cf; else sh.conditionalFormats.push(cf); }); renderRight();
  } }, { label: T('cancel') }] });
}
function dvDialog(idx) {
  const sh = KXL.cur(); const old = idx != null ? sh.validations[idx] : null;
  const range = el('input', { type: 'text', dir: 'ltr', value: old ? old.range : KXL.rangeStr(S.sel) });
  const type = sel([['list', T('vList')], ['whole', T('vWhole')], ['decimal', T('vDecimal')], ['date', T('vDate')], ['textLength', T('vLen')], ['custom', T('vCustom')]], old ? old.type : 'list', () => sync());
  const op = sel(OPS.filter(o => !['notEqual'].includes(o) || true).map(o => [o, opLabel(o)]), old ? old.operator || 'between' : 'between', () => sync());
  const f1 = el('input', { type: 'text', dir: 'ltr', value: old ? old.formula1 || '' : '' }), f2 = el('input', { type: 'text', dir: 'ltr', value: old ? old.formula2 || '' : '' });
  const err = el('input', { type: 'text', value: old ? old.error || '' : '' }), prompt = el('input', { type: 'text', value: old ? old.prompt || '' : '' });
  const rOp = frow(T('operator'), op), rF2 = frow(T('value2'), f2); const hint = el('div', { class: 'help' });
  function sync() { const t = type.value; rOp.style.display = ['list', 'custom'].includes(t) ? 'none' : ''; rF2.style.display = (!['list', 'custom'].includes(t) && (op.value === 'between' || op.value === 'notBetween')) ? '' : 'none'; hint.textContent = t === 'list' ? T('vListHelp') : ''; }
  const body = el('div', {}, frow(T('rangeLbl'), range), frow(T('type'), type), rOp, frow(T('value1'), f1), rF2, hint, frow(T('errMsg'), err), frow('Prompt', prompt));
  sync();
  UI.modal({ title: T('validation'), body, buttons: [{ label: old ? T('ok') : T('add'), primary: true, onclick: () => {
    if (!KXL.parseRange(range.value) || !f1.value.trim()) { KXL.toast(T('rangeLbl'), 'err'); return false; }
    const v = { range: range.value.trim().toUpperCase(), type: type.value, operator: op.value, formula1: f1.value, formula2: f2.value || undefined, allowBlank: true, error: err.value || undefined, prompt: prompt.value || undefined };
    KXL.mutate(() => { if (old) sh.validations[idx] = v; else sh.validations.push(v); }); renderRight();
  } }, { label: T('cancel') }] });
}

function chartDialog(idx) {
  const sh = KXL.cur(); const old = idx != null ? sh.charts[idx] : null;
  const type = sel([['column', T('tColumn')], ['bar', T('tBar')], ['line', T('tLine')], ['area', T('tArea')], ['pie', T('tPie')]], old ? old.type : 'column', () => {});
  const title = el('input', { type: 'text', value: old ? old.title || '' : '' });
  const row = el('input', { type: 'number', value: old ? old.row : S.sel.ar }), col = el('input', { type: 'number', value: old ? old.col : S.sel.ac + 2 });
  const w = el('input', { type: 'number', value: old ? old.width : 480 }), h = el('input', { type: 'number', value: old ? old.height : 288 });
  const cats = el('input', { type: 'text', dir: 'ltr', value: old ? old.categories || '' : '', placeholder: 'B9:B9' });
  const stacked = el('input', { type: 'checkbox' }); stacked.checked = !!(old && old.stacked);
  const legend = el('input', { type: 'checkbox' }); legend.checked = old ? old.showLegend !== false : true;
  const labels = el('input', { type: 'checkbox' }); labels.checked = !!(old && old.showDataLabels);
  const list = el('div');
  const rows = [];
  const addRow = (s) => {
    const name = el('input', { type: 'text', placeholder: T('seriesName'), value: s.name || '', style: 'flex:1;min-width:0' });
    const vals = el('input', { type: 'text', dir: 'ltr', placeholder: T('seriesValues'), value: s.values || '', style: 'flex:1.3;min-width:0' });
    const c = el('input', { type: 'color', value: s.color || '#2e75b6', style: 'width:36px;height:30px' });
    const r = { name, vals, c }; rows.push(r);
    const line = el('div', { style: 'display:flex;gap:6px;margin-bottom:6px;align-items:center' }, name, vals, c, el('button', { class: 'iconbtn', onclick: () => { rows.splice(rows.indexOf(r), 1); line.remove(); } }, '\u2715'));
    list.append(line);
  };
  ((old && old.series && old.series.length) ? old.series : [{}]).forEach(addRow);
  const body = el('div', {}, el('div', { class: 'help' }, T('chartHelp')),
    el('div', { class: 'two' }, frow(T('chartType'), type), frow(T('chartTitle'), title)),
    el('div', { class: 'two' }, frow('Row', row), frow('Col', col)), el('div', { class: 'two' }, frow('W', w), frow('H', h)),
    frow(T('categories'), cats), el('div', { class: 'lbl2' }, T('series')), list,
    el('button', { class: 'btn', onclick: () => addRow({}) }, '+ ' + T('addSeries')),
    el('div', { style: 'margin-top:10px' }, el('label', { class: 'chk' }, stacked, T('stacked')), el('label', { class: 'chk' }, legend, T('legend')), el('label', { class: 'chk' }, labels, T('dataLabels'))));
  UI.modal({ title: T('addChart'), body, buttons: [{ label: old ? T('ok') : T('add'), primary: true, onclick: () => {
    const series = rows.filter(r => KXL.parseRange(r.vals.value)).map(r => { const s = { values: r.vals.value.trim().toUpperCase(), color: r.c.value }; if (r.name.value.trim()) { if (KXL.parseRange(r.name.value)) s.nameCell = r.name.value.trim().toUpperCase(); else s.name = r.name.value.trim(); } return s; });
    if (!series.length) { KXL.toast(T('seriesValues'), 'err'); return false; }
    const ch = { name: old ? old.name : 'Chart' + (sh.charts.length + 1), type: type.value, title: title.value || undefined, row: Math.max(1, parseInt(row.value) || 1), col: Math.max(1, parseInt(col.value) || 1), width: Math.max(120, parseInt(w.value) || 480), height: Math.max(80, parseInt(h.value) || 288), categories: cats.value.trim() ? cats.value.trim().toUpperCase() : undefined, series, stacked: stacked.checked || undefined, showLegend: legend.checked, showDataLabels: labels.checked || undefined };
    KXL.mutate(() => { if (old) sh.charts[idx] = ch; else sh.charts.push(ch); }); G.render(); S.panelTab = 'images'; renderRight();
  } }, { label: T('cancel') }] });
}

function tabPage(body) {
  const sh = KXL.cur(); const pg = sh.page = sh.page || {}; const set = (k, v) => KXL.mutate(() => { pg[k] = v; });
  body.append(frow(T('orientation'), sel([['portrait', T('portrait')], ['landscape', T('landscape')]], pg.orientation || 'portrait', (v) => set('orientation', v))));
  body.append(frow(T('paper'), sel([[9, 'A4'], [8, 'A3'], [11, 'A5'], [1, 'Letter'], [5, 'Legal']], pg.paperSize || 9, (v) => set('paperSize', parseInt(v)))));
  body.append(chk(T('fitWidth'), pg.fitToWidth, (v) => set('fitToWidth', v)));
  body.append(chk(T('center2'), pg.centerHorizontally, (v) => set('centerHorizontally', v)));
  body.append(el('div', { class: 'two' }, ...['marginTop', 'marginBottom', 'marginLeft', 'marginRight'].map(k => frow(T('margins') + ' · ' + k.replace('margin', ''), inp(pg[k] != null ? pg[k] : (k.match(/Top|Bottom/) ? 0.75 : 0.5), (v) => set(k, parseFloat(v) || 0), { type: 'number', step: '0.05' })))));
  body.append(frow(T('header'), inp(pg.header || '', (v) => set('header', v || undefined), { dir: 'ltr', placeholder: '&C&"Tahoma,Bold"Title' })));
  body.append(frow(T('footer'), inp(pg.footer || '', (v) => set('footer', v || undefined), { dir: 'ltr', placeholder: '&CPage &P of &N' })));
  body.append(frow(T('printTitles'), inp(pg.printTitleRows || '', (v) => set('printTitleRows', v || undefined), { dir: 'ltr', placeholder: '1:3' })));
  body.append(frow(T('printArea'), inp(pg.printArea || '', (v) => set('printArea', v || undefined), { dir: 'ltr', placeholder: 'A1:F40' })));
  body.append(frow(T('watermark'), inp(pg.watermark || '', (v) => set('watermark', v || undefined), { placeholder: 'DRAFT' })));
  body.append(el('div', { class: 'help', dir: 'ltr', style: 'text-align:left' }, '&P page · &N pages · &D date · &T time · &F file · &L &C &R sections'));
}

/* ---- images */
function pickImage() {
  const f = el('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif' });
  f.addEventListener('change', () => { const file = f.files[0]; if (!file) return; const rd = new FileReader(); rd.onload = () => addImage(rd.result, file.name); rd.readAsDataURL(file); });
  f.click();
}
function addImage(dataUrl, name) {
  const img = new Image();
  img.onload = () => {
    let w = img.naturalWidth, h = img.naturalHeight; const maxW = 1200;
    let url = dataUrl;
    if (w > maxW) { const k = maxW / w; const cv = document.createElement('canvas'); cv.width = Math.round(w * k); cv.height = Math.round(h * k); cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height); url = cv.toDataURL(dataUrl.startsWith('data:image/jpeg') ? 'image/jpeg' : 'image/png', .9); w = cv.width; h = cv.height; }
    const dw = Math.min(w, 220), dh = Math.round(h * dw / w);
    const m = /^data:image\/(\w+);base64,(.*)$/.exec(url); if (!m) return;
    const ext = m[1] === 'jpg' ? 'jpeg' : m[1];
    KXL.mutate(() => { const sh = KXL.cur(); const id = 'b' + (++S.blobSeq); S.blobs[id] = m[2]; sh.images.push({ name: name || 'Image', row: S.sel.ar, col: S.sel.ac, width: dw, height: dh, ext, dataRef: id }); });
    G.render(); S.panelTab = 'images'; renderRight();
  };
  img.src = dataUrl;
}
function tabImages(body) {
  const sh = KXL.cur();
  body.append(el('div', { class: 'sec' }, el('h4', {}, T('charts')), el('button', { class: 'btn primary', onclick: () => chartDialog() }, '+ ' + T('addChart'))));
  if (!sh.charts.length) body.append(el('div', { class: 'muted', style: 'margin-bottom:12px' }, T('noCharts')));
  sh.charts.forEach((ch, i) => body.append(el('div', { class: 'card' }, el('div', { class: 'hd' }, '\u{1F4CA} ' + (ch.title || ch.name || ch.type), el('span', { class: 'pill' }, ch.type),
    el('button', { class: 'iconbtn x', title: T('edit'), onclick: () => chartDialog(i) }, '\u270E'), el('button', { class: 'iconbtn', title: T('del'), onclick: () => { KXL.mutate(() => { sh.charts.splice(i, 1); }); G.render(); renderRight(); } }, '\u2715')),
    el('div', { class: 'muted', dir: 'ltr', style: 'text-align:start' }, (ch.categories || '') + ' | ' + (ch.series || []).map(s => s.values).join(', ')))));
  body.append(el('hr', { style: 'border:0;border-top:1px solid var(--line2);margin:12px 0' }));
  body.append(el('div', { class: 'sec' }, el('h4', {}, T('images'))));
  body.append(el('button', { class: 'btn primary', style: 'margin-bottom:10px', onclick: pickImage }, '+ ' + T('pickImage')));
  body.append(el('div', { class: 'help' }, T('imgDynamicHelp')));
  if (!sh.images.length) body.append(el('div', { class: 'empty' }, '—'));
  sh.images.forEach((im, i) => {
    const prev = S.blobs[im.dataRef] ? el('img', { src: 'data:image/' + im.ext + ';base64,' + S.blobs[im.dataRef], style: 'max-width:100%;max-height:70px;display:block;margin-bottom:6px;border:1px solid var(--line2)' }) : null;
    const upd = (k, v) => { KXL.mutate(() => { im[k] = v; }); G.render(); };
    body.append(el('div', { class: 'card' }, el('div', { class: 'hd' }, im.name || 'Image', el('button', { class: 'iconbtn x', onclick: () => { KXL.mutate(() => { sh.images.splice(i, 1); }); G.render(); renderRight(); } }, '✕')), prev,
      el('div', { class: 'two' }, frow('Row', inp(im.row, (v) => upd('row', Math.max(1, parseInt(v) || 1)), { type: 'number' })), frow('Col', inp(im.col, (v) => upd('col', Math.max(1, parseInt(v) || 1)), { type: 'number' }))),
      el('div', { class: 'two' }, frow('W', inp(im.width, (v) => upd('width', Math.max(4, parseInt(v) || 60)), { type: 'number' })), frow('H', inp(im.height, (v) => upd('height', Math.max(4, parseInt(v) || 30)), { type: 'number' }))),
      frow(T('imgBinding'), inp(im.binding || '', (v) => upd('binding', v.trim() || undefined), { dir: 'ltr', placeholder: 'Company.Logo' }))));
  });
}

/* ------------------------------------------------------------------ context menu */
KXL.fn.showContextMenu = function (x, y) {
  const m = $('#ctxmenu'); m.innerHTML = '';
  const s = S.sel; const multi = s.r1 !== s.r2 || s.c1 !== s.c2; const sh = KXL.cur();
  const add = (label, run, dis) => m.append(el('button', { class: 'mi' + (dis ? ' dis' : ''), onclick: () => { m.style.display = 'none'; run(); } }, label));
  const sep = () => m.append(el('div', { class: 'msep' }));
  add(T('menuCut'), () => { G.copy(null, true); }); add(T('menuCopy'), () => G.copy(null, false));
  add(T('menuPaste'), async () => { try { const t = await navigator.clipboard.readText(); G.paste(t); } catch (e) { if (S.clip) G.paste(S.clip.text); else KXL.toast('Ctrl+V', 'err'); } });
  sep(); add(T('clearCells'), () => G.clearContents()); add(T('clearFmt'), () => G.clearFormats());
  sep(); add(T('merge'), () => G.merge(), !multi || rangeIsMerge()); add(T('unmerge'), () => G.unmerge(), !KXL.mergeAt(s.ar, s.ac));
  sep(); add(T('insRowAbove'), () => insertRows(false)); add(T('insRowBelow'), () => insertRows(true)); add(T('delRow'), deleteRows);
  add(T('insColBefore'), () => insertCols(false)); add(T('insColAfter'), () => insertCols(true)); add(T('delCol'), deleteCols);
  sep(); add(T('addBand') + ' (' + s.r1 + (s.r2 > s.r1 ? '–' + s.r2 : '') + ')', () => bandDialog());
  const bi = sh.bands.findIndex(b => s.ar >= b.firstRow && s.ar <= b.lastRow); if (bi >= 0) add(T('del') + ' ' + T('bandLbl') + ' ' + (sh.bands[bi].name || ''), () => { KXL.mutate(() => { sh.bands.splice(bi, 1); }); G.render(); renderRight(); });
  m.style.display = 'block';
  const w = m.offsetWidth, h = m.offsetHeight; m.style.left = Math.min(x, innerWidth - w - 6) + 'px'; m.style.top = Math.min(y, innerHeight - h - 6) + 'px';
};

/* ------------------------------------------------------------------ file operations */
const NAME_RX = /^[\p{L}\p{N}][\p{L}\p{N} _\-.]{0,98}$/u;
const validName = (n) => NAME_RX.test(n || '') && !n.includes('..') && !n.endsWith('.') && !n.endsWith(' ');

async function onSave() {
  if (S.cfg.readOnly) { KXL.toast(T('readOnly'), 'err'); return; }
  if (S.editing) G.commitEdit();
  let name = (S.tpl.name || '').trim();
  if (!validName(name)) {
    const v = await UI.prompt(T('templateName'), name === T('untitled') ? '' : name, T('invalidName'));
    if (!v) return; name = v.trim(); if (!validName(name)) { KXL.toast(T('invalidName'), 'err'); return; }
  }
  try {
    S.tpl.name = name; $('#tpl-name').value = name;
    await KXL.api('PUT', 'templates/' + encodeURIComponent(name), KXL.serialize());
    S.savedName = name; S.dirty = false; updateDirty(); KXL.toast(T('saved'), 'ok'); clearDraft();
    history.replaceState(null, '', '?t=' + encodeURIComponent(name));
  } catch (e) { KXL.toast(T('saveFailed') + ': ' + e.message, 'err'); }
}
async function guardDiscard() { return !S.dirty || await UI.confirm(T('confirmDiscard'), true); }

async function onNew() {
  if (!(await guardDiscard())) return;
  const box = el('div', { class: 'startgrid' });
  let m;
  const card = (title, sub, run) => box.append(el('div', { class: 'startcard', onclick: async () => { m.close(); await run(); } }, el('b', {}, title), el('span', { class: 'muted' }, sub || '')));
  card(T('blankRtl'), 'A1…', () => KXL.load({ name: '', title: '', sheets: [{ name: T('sheet') + '1', rtl: true }], styles: {} }));
  card(T('blank'), 'A1…', () => KXL.load({ name: '', title: '', sheets: [{ name: T('sheet') + '1', rtl: false }], styles: {} }));
  for (const st of (S.cfg.starters || [])) card(T('starter_' + st.key) === 'starter_' + st.key ? st.title : T('starter_' + st.key), st.key, async () => { const t = await KXL.api('GET', 'starters/' + st.key); t.name = ''; if (KXL.RTL) t.sheets.forEach(s => { s.rtl = true; }); KXL.load(t); });
  m = UI.modal({ title: T('newTitle'), body: box });
}
async function onOpen() {
  if (!(await guardDiscard())) return;
  let list = []; try { list = await KXL.api('GET', 'templates'); } catch (e) { KXL.toast(e.message, 'err'); return; }
  const host = el('div', { class: 'tlist' });
  let m;
  const draw = () => {
    host.innerHTML = ''; if (!list.length) host.append(el('div', { class: 'empty' }, T('noTemplates')));
    for (const t of list) host.append(el('div', { class: 'row', onclick: async () => { m.close(); await openByName(t.name); } }, el('span', { class: 'nm' }, t.name), el('small', {}, new Date(t.modified).toLocaleString()),
      S.cfg.readOnly ? null : el('button', { class: 'iconbtn', title: T('del'), onclick: async (e) => { e.stopPropagation(); if (await UI.confirm(T('confirmDel') + ' ' + t.name, true)) { await KXL.api('DELETE', 'templates/' + encodeURIComponent(t.name)); list = list.filter(x => x.name !== t.name); draw(); } } }, '✕')));
  };
  draw(); m = UI.modal({ title: T('openTitle'), body: host });
}
async function openByName(name) {
  try { const t = await KXL.api('GET', 'templates/' + encodeURIComponent(name)); KXL.load(t, name); history.replaceState(null, '', '?t=' + encodeURIComponent(name)); }
  catch (e) { KXL.toast(e.message, 'err'); }
}
KXL.fn.openByName = openByName;

async function onPreview() {
  if (S.editing) G.commitEdit();
  const host = el('div', { id: 'preview-host' }, T('loading'));
  const m = UI.modal({ title: T('previewTitle'), wide: true, body: host, buttons: [{ label: T('download'), primary: true, onclick: () => { onDownload(); return false; } }, { label: T('close') }] });
  m.body.style.cssText = 'display:flex;flex-direction:column;padding:12px;min-height:0';
  try {
    const r = await KXL.api('POST', 'preview', { template: KXL.serialize() });
    host.innerHTML = '';
    host.append(el('div', { class: r.warnings.length ? 'warnbox' : 'warnbox ok' }, r.warnings.length ? el('div', {}, el('b', {}, T('warnings') + ': '), ...r.warnings.map(w => el('div', {}, '• ' + w))) : T('noWarn')));
    host.append(el('div', { html: r.html }));
  } catch (e) { host.innerHTML = ''; host.append(el('div', { class: 'warnbox' }, e.message)); }
}
async function onDownload() {
  if (S.editing) G.commitEdit();
  try {
    const res = await KXL.api('POST', 'download', { template: KXL.serialize() }, true);
    const blob = await res.blob(); const a = el('a', { href: URL.createObjectURL(blob), download: (S.tpl.name || 'report') + '.xlsx' }); document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  } catch (e) { KXL.toast(e.message, 'err'); }
}
async function onPdf() {
  if (S.editing) G.commitEdit();
  try {
    const res = await KXL.api('POST', 'pdf', { template: KXL.serialize() }, true);
    const blob = await res.blob(); const a = el('a', { href: URL.createObjectURL(blob), download: (S.tpl.name || 'report') + '.pdf' }); document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  } catch (e) { KXL.toast(T('pdfFail') + ': ' + e.message, 'err'); }
}
async function importXlsx() {
  if (!(await guardDiscard())) return;
  const f = el('input', { type: 'file', accept: '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  f.addEventListener('change', () => {
    const file = f.files[0]; if (!file) return;
    const rd = new FileReader();
    rd.onload = async () => {
      try {
        const b64 = String(rd.result).split(',')[1];
        const r = await KXL.api('POST', 'import', { name: file.name, data: b64 });
        KXL.load(r.template, null); S.dirty = true; KXL.fn.changed(true);
        KXL.toast(T('importDone'), 'ok');
        if (r.warnings && r.warnings.length) setTimeout(() => KXL.toast(r.warnings.slice(0, 3).join(' | '), 'err'), 600);
      } catch (e) { KXL.toast(T('importFail') + ': ' + e.message, 'err'); }
    };
    rd.readAsDataURL(file);
  });
  f.click();
}
function exportJson() {
  const blob = new Blob([JSON.stringify(KXL.serialize(), null, 2)], { type: 'application/json' });
  const a = el('a', { href: URL.createObjectURL(blob), download: (S.tpl.name || 'template') + '.kxl.json' }); document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}
async function importJson() {
  if (!(await guardDiscard())) return;
  const f = el('input', { type: 'file', accept: '.json,application/json' });
  f.addEventListener('change', () => { const file = f.files[0]; if (!file) return; const rd = new FileReader(); rd.onload = () => { try { const t = JSON.parse(rd.result); if (!t.sheets) throw new Error('not a template'); KXL.load(t, null); S.dirty = true; updateDirty(); } catch (e) { KXL.toast(T('invalidJson') + ': ' + e.message, 'err'); } }; rd.readAsText(file); });
  f.click();
}

/* ---- draft autosave (browser only; best effort) */
let draftTimer = null;
function scheduleDraft() { clearTimeout(draftTimer); draftTimer = setTimeout(() => { try { localStorage.setItem('kxl-draft:' + location.pathname, JSON.stringify({ at: Date.now(), name: S.tpl.name, data: KXL.serialize() })); } catch (e) { /* storage unavailable */ } }, 1200); }
function clearDraft() { try { localStorage.removeItem('kxl-draft:' + location.pathname); } catch (e) { } }
KXL.fn.readDraft = function () { try { const d = JSON.parse(localStorage.getItem('kxl-draft:' + location.pathname) || 'null'); return d && d.data ? d : null; } catch (e) { return null; } };

function updateDirty() { const d = $('#dirty'); d.className = S.dirty ? 'on' : ''; d.title = S.dirty ? T('unsaved') : ''; document.title = (S.tpl && S.tpl.name ? S.tpl.name + ' - ' : '') + ($('#app-title').textContent); }

/* ------------------------------------------------------------------ hooks */
KXL.fn.hydrate = function (sh) { KXL.fn.hydrateOne(sh); sh._gr = 0; sh._gc = 0; };
KXL.fn.selectionChanged = function () { updateFormulaBar(); refreshToolbar(); if (S.panelTab === 'cell') renderRight(); };
KXL.fn.changed = function (full) { updateDirty(); renderSheetTabs(); updateFormulaBar(); refreshToolbar(); renderRight(); scheduleDraft(); if (full) { renderLeft(); G.render(); } };
KXL.fn.afterLoad = function () {
  $('#tpl-name').value = S.tpl.name || ''; $('#tpl-name').placeholder = T('untitled');
  renderSheetTabs(); renderLeft(); renderRight(); G.render(); updateFormulaBar(); refreshToolbar(); updateDirty();
  $('#gridscroll').scrollTop = 0; $('#gridscroll').scrollLeft = 0;
  if (document.activeElement === document.body) $('#gridwrap').focus({ preventScroll: true });
};

/* ------------------------------------------------------------------ global shortcuts */
document.addEventListener('keydown', (e) => {
  const mod = e.ctrlKey || e.metaKey; if (!mod) { if (e.key === 'Escape') { UI.closePopover(); $('#ctxmenu').style.display = 'none'; } return; }
  const a = document.activeElement; const typing = a && (a.tagName === 'INPUT' || a.tagName === 'TEXTAREA' || a.tagName === 'SELECT') && a.id !== 'cell-editor';
  const k = e.key.toLowerCase();
  if (k === 's') { e.preventDefault(); onSave(); return; }
  if (typing || S.editing) return;
  if (k === 'z') { e.preventDefault(); KXL.undoOp(); } else if (k === 'y') { e.preventDefault(); KXL.redoOp(); }
  else if (k === 'b') { e.preventDefault(); G.applyStyle({ bold: !styleNow().bold }); } else if (k === 'i') { e.preventDefault(); G.applyStyle({ italic: !styleNow().italic }); } else if (k === 'u') { e.preventDefault(); G.applyStyle({ underline: !styleNow().underline }); }
  else if (k === 'a') { e.preventDefault(); G.select(1, 1, KXL.cur()._gr, KXL.cur()._gc); }
});
window.addEventListener('beforeunload', (e) => { if (S.dirty) { e.preventDefault(); e.returnValue = ''; } });

/* ------------------------------------------------------------------ init UI pieces */
UI.init = function () {
  buildTop(); buildRibbon(); wireFormulaBar(); G.init();
  $('#app-title').textContent = (S.cfg && S.cfg.title) || T('appName');
  if (S.cfg && S.cfg.readOnly) { $('#top-actions .btn.primary').disabled = true; KXL.toast(T('readOnly')); }
};
UI.renderLeft = renderLeft; UI.openByName = openByName; UI.onNew = onNew;
})();
