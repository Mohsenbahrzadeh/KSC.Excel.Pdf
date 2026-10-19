/* KSC.Excel designer - bootstrap. */
(async function () {
'use strict';
const KXL = window.KXL, S = KXL.S, T = KXL.T, el = KXL.el;
try { S.cfg = await KXL.api('GET', 'config'); }
catch (e) { document.body.innerHTML = '<pre style="padding:20px;color:#c0392b">' + KXL.esc(e.message) + '</pre>'; return; }

KXL.ui.init();
const blank = () => KXL.load({ name: '', title: '', sheets: [{ name: T('sheet') + '1', rtl: KXL.RTL }], styles: {} });
const q = new URLSearchParams(location.search).get('t');
blank();
if (q) await KXL.fn.openByName(q);

const draft = KXL.fn.readDraft();
if (draft && (!q || draft.name === q)) {
  const when = new Date(draft.at).toLocaleString();
  KXL.ui.modal({
    title: draft.name ? draft.name + ' — ' + when : when,
    body: el('div', {}, T('restoreAsk')),
    buttons: [{ label: T('restore'), primary: true, onclick: () => { KXL.load(draft.data, draft.name || null); S.dirty = true; KXL.fn.changed(true); } },
              { label: T('discard'), onclick: () => { try { localStorage.removeItem('kxl-draft:' + location.pathname); } catch (e) { } } }]
  });
}
window.addEventListener('error', (e) => KXL.toast(String(e.message || e), 'err'));
})();
