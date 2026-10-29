(() => {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const pad = (n) => String(n).padStart(2, '0');

  /* ---------- Hero: switch sides ---------- */
  const segButtons = $$('.seg__btn');
  segButtons.forEach((btn) => {
    btn.addEventListener('click', () => {
      const role = btn.dataset.role;
      segButtons.forEach((b) => b.setAttribute('aria-pressed', String(b === btn)));
      $$('.app__view').forEach((v) => { v.hidden = v.dataset.view !== role; });
    });
  });

  /* ---------- Hero: the real time, and clocking in ---------- */
  const timeEl = $('#clock-time');
  const dateEl = $('#app-date');
  const clockBtn = $('#clock-btn');
  const statusEl = $('#clock-status');
  let clockedAt = null;

  const greet = () => {
    const h = new Date().getHours();
    const hello = $('.app__view[data-view="employee"] .app__hello');
    if (hello) hello.textContent = (h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening') + ', Priya';
  };

  const tick = () => {
    const now = new Date();
    if (timeEl) timeEl.textContent = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
    if (clockedAt && statusEl) {
      const s = Math.floor((now - clockedAt) / 1000);
      statusEl.innerHTML = `On shift for <b class="num">${pad(Math.floor(s / 3600))}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}</b> · clocked in at ${pad(clockedAt.getHours())}:${pad(clockedAt.getMinutes())}`;
    }
  };
  if (dateEl) dateEl.textContent = new Date().toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long' });
  greet();
  tick();
  setInterval(tick, 1000);

  clockBtn?.addEventListener('click', () => {
    if (clockedAt) {
      const mins = Math.max(1, Math.round((Date.now() - clockedAt) / 60000));
      clockedAt = null;
      clockBtn.textContent = 'Clock in';
      clockBtn.setAttribute('aria-pressed', 'false');
      statusEl.textContent = `Clocked out. ${mins} min added to this week’s timesheet.`;
    } else {
      clockedAt = new Date();
      clockBtn.textContent = 'Clock out';
      clockBtn.setAttribute('aria-pressed', 'true');
      tick();
    }
  });

  /* ---------- Hero: employer approvals ---------- */
  const approveBtn = $('#approve-btn');
  approveBtn?.addEventListener('click', () => {
    approveBtn.disabled = true;
    approveBtn.textContent = 'Approved';
    $('#approve-text').textContent = 'All caught up. Payroll has the hours.';
  });

  /* ---------- Features: command list ---------- */
  const input = $('#cmd-input');
  const rows = $$('.cmd__row');
  const empty = $('#cmd-empty');
  const visible = () => rows.filter((r) => !r.hidden);
  const select = (row, focus = false) => {
    if (!row) return;
    rows.forEach((r) => { const on = r === row; r.setAttribute('aria-selected', String(on)); r.tabIndex = on ? 0 : -1; });
    if (focus) row.focus({ preventScroll: false });
  };
  rows.forEach((row) => {
    row.addEventListener('click', () => select(row));
    row.addEventListener('mouseenter', () => { if (matchMedia('(hover: hover)').matches) select(row); });
    row.addEventListener('keydown', (e) => {
      const list = visible();
      const i = list.indexOf(row);
      if (e.key === 'ArrowDown') { e.preventDefault(); select(list[Math.min(i + 1, list.length - 1)], true); }
      if (e.key === 'ArrowUp') { e.preventDefault(); if (i === 0) input.focus(); else select(list[i - 1], true); }
    });
  });
  input?.addEventListener('input', () => {
    const q = input.value.trim().toLowerCase();
    rows.forEach((r) => { r.hidden = q !== '' && !r.textContent.toLowerCase().includes(q); });
    const list = visible();
    empty.hidden = list.length > 0;
    if (list.length && !list.some((r) => r.getAttribute('aria-selected') === 'true')) select(list[0]);
  });
  input?.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') { e.preventDefault(); select(visible()[0], true); }
  });
  // Single-letter shortcuts while the list is in view and nothing is being typed
  document.addEventListener('keydown', (e) => {
    if (e.metaKey || e.ctrlKey || e.altKey || /^(INPUT|SELECT|TEXTAREA)$/.test(document.activeElement?.tagName)) return;
    const list = $('#cmd-list');
    const r = list?.getBoundingClientRect();
    if (!r || r.bottom < 0 || r.top > innerHeight) return;
    const hit = rows.find((row) => row.querySelector('kbd')?.textContent.toLowerCase() === e.key.toLowerCase());
    if (hit && !hit.hidden) select(hit);
  });

  /* ---------- How it works: the request changes with each statement ---------- */
  const req = $('#req');
  const steps = $$('.how__step');
  if (req && steps.length && 'IntersectionObserver' in window) {
    const so = new IntersectionObserver((entries) => {
      entries.forEach((e) => {
        if (!e.isIntersecting) return;
        steps.forEach((s) => s.classList.toggle('is-active', s === e.target));
        req.dataset.state = e.target.dataset.step;
      });
    }, { rootMargin: '-45% 0px -45% 0px' });
    steps.forEach((s) => so.observe(s));
  } else {
    steps.forEach((s) => s.classList.add('is-active'));
  }

  /* ---------- Start: the form opens a filled-in email ---------- */
  $('#start-form')?.addEventListener('submit', (e) => {
    e.preventDefault();
    const f = e.currentTarget;
    const company = f.company.value.trim() || 'my company';
    const side = f.querySelector('input[name="side"]:checked')?.value || 'an employer';
    const size = f.size.value;
    const subject = `Workzenly demo for ${company}`;
    const body = `Hello Workzenly,\n\nI would like a demo.\n\nCompany: ${company}\nI am: ${side}\nTeam size: ${size}\n\nThank you.`;
    location.href = `mailto:hello@workzenly.com?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  });

  /* ---------- Reveals, staggered inside each section ---------- */
  $$('main section, footer').forEach((sec) => {
    $$('.reveal', sec).forEach((el, i) => el.style.setProperty('--d', `${Math.min(i, 6) * 80}ms`));
  });
  const reveals = $$('.reveal');
  if ('IntersectionObserver' in window) {
    const io = new IntersectionObserver((entries) => {
      entries.forEach((e) => { if (e.isIntersecting) { e.target.classList.add('is-in'); io.unobserve(e.target); } });
    }, { rootMargin: '0px 0px -10% 0px', threshold: 0.08 });
    reveals.forEach((el) => io.observe(el));
  } else {
    reveals.forEach((el) => el.classList.add('is-in'));
  }

  const y = $('#year');
  if (y) y.textContent = new Date().getFullYear();
})();
