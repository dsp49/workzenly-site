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
    if (hello) hello.textContent = (h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening') + ', Maya';
  };

  const tick = () => {
    const now = new Date();
    if (timeEl) timeEl.textContent = `${now.getHours() % 12 || 12}:${pad(now.getMinutes())}:${pad(now.getSeconds())} ${now.getHours() < 12 ? 'AM' : 'PM'}`;
    if (clockedAt && statusEl) {
      const s = Math.floor((now - clockedAt) / 1000);
      statusEl.innerHTML = `On shift for <b class="num">${pad(Math.floor(s / 3600))}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}</b> · clocked in at ${clockedAt.getHours() % 12 || 12}:${pad(clockedAt.getMinutes())} ${clockedAt.getHours() < 12 ? 'AM' : 'PM'}`;
    }
  };
  if (dateEl) dateEl.textContent = new Date().toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
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

  /* ---------- How it works: the request changes with each statement ---------- */
  // Wide screens: a step takes over when it reaches the middle of the window,
  // beside the card. Phones: the card is pinned above the text, so a step takes
  // over as its headline comes into view below the card, not once it has
  // already scrolled underneath it.
  const req = $('#req');
  const steps = $$('.how__step');
  const phone = matchMedia('(max-width: 760px)');
  if (req && steps.length) {
    let current = null;
    const updateStep = () => {
      const line = phone.matches ? innerHeight - 48 : innerHeight * 0.55;
      let active = steps[0];
      steps.forEach((s) => {
        const edge = phone.matches ? $('.how__line', s) : s;
        if (edge.getBoundingClientRect().top <= line) active = s;
      });
      if (active === current) return;
      current = active;
      steps.forEach((s) => s.classList.toggle('is-active', s === active));
      req.dataset.state = active.dataset.step;
    };
    updateStep();
    addEventListener('scroll', updateStep, { passive: true });
    addEventListener('resize', updateStep);
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
    location.href = `mailto:info@workzenly.com?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
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
