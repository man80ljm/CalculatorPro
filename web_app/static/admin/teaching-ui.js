window.TeachingAdmin = { timetable: { normalized: value => String(value || "").normalize("NFKC").toLowerCase().replace(/[^\p{L}\p{N}]/gu, "") } };
(function (P) {
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
  let timer, priorFocus;
  function toast(message) { const el = document.getElementById('toast'); el.textContent = message; el.hidden = false; clearTimeout(timer); timer = setTimeout(() => { el.hidden = true; }, 4000); }
  function show({ title, subtitle = '', eyebrow = '', body, foot = '', wide = false }) {
    const dialog = document.getElementById('appDialog');
    dialog.classList.toggle('wide', wide);
    document.getElementById('dialogContent').innerHTML = '<div class="dialog-head"><div>' + (eyebrow ? '<span class="eyebrow">' + esc(eyebrow) + '</span>' : '') + '<h2 id="dialogTitle">' + esc(title) + '</h2>' + (subtitle ? '<p>' + esc(subtitle) + '</p>' : '') + '</div><button class="close-button" data-action="close-dialog" aria-label="关闭弹窗">×</button></div><div class="dialog-body">' + body + '</div>' + (foot ? '<div class="dialog-foot">' + foot + '</div>' : '');
    if (!dialog.open) { priorFocus = document.activeElement; dialog.showModal(); }
  }
  function close() { document.getElementById('appDialog').close(); }
  function button(text, action, extra = '', secondary = false) { return '<button type="button"' + (secondary ? ' class="secondary"' : '') + ' data-action="' + action + '" ' + extra + '>' + esc(text) + '</button>'; }
  function options(items, value) { return items.map(i => '<option value="' + esc(i.id) + '"' + (i.id === value ? ' selected' : '') + '>' + esc(i.label || i.name) + '</option>').join(''); }
  document.getElementById('appDialog').addEventListener('close', () => { if (priorFocus?.isConnected) priorFocus.focus(); });
  document.addEventListener('click', event => { const control = event.target.closest('[data-action]'); if (control?.dataset.action === 'close-dialog') close(); });
  P.ui = { esc, toast, show, close, button, options };
})(window.TeachingAdmin);
