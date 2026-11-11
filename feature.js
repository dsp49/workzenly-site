(() => {
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));

  /* ---------- Product tabs (WAI-ARIA tabs pattern) ---------- */
  $$('[data-tabs]').forEach((root) => {
    const tabs = $$('[role="tab"]', root);
    const show = (tab, focus = false) => {
      tabs.forEach((t) => {
        const on = t === tab;
        t.setAttribute('aria-selected', String(on));
        t.tabIndex = on ? 0 : -1;
        document.getElementById(t.getAttribute('aria-controls')).hidden = !on;
      });
      if (focus) tab.focus();
    };
    tabs.forEach((tab, i) => {
      tab.addEventListener('click', () => show(tab));
      tab.addEventListener('keydown', (e) => {
        const next = { ArrowRight: 1, ArrowDown: 1, ArrowLeft: -1, ArrowUp: -1 }[e.key];
        if (next) { e.preventDefault(); show(tabs[(i + next + tabs.length) % tabs.length], true); }
        if (e.key === 'Home') { e.preventDefault(); show(tabs[0], true); }
        if (e.key === 'End') { e.preventDefault(); show(tabs[tabs.length - 1], true); }
      });
    });
  });

  /* ---------- Buttons inside product mockups: one click, one visible result ---------- */
  $$('[data-act]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.target && document.getElementById(btn.dataset.target);
      if (target && btn.dataset.text) target.textContent = btn.dataset.text;
      if (btn.dataset.pill) {
        const pill = document.getElementById(btn.dataset.pill);
        if (pill) { pill.className = 'ui-pill ui-pill--ok'; pill.textContent = btn.dataset.pillText || 'Approved'; }
      }
      btn.textContent = btn.dataset.done || 'Done';
      btn.disabled = true;
    });
  });

  const y = document.getElementById('year');
  if (y) y.textContent = new Date().getFullYear();
})();
