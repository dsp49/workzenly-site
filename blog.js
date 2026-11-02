(() => {
  // Copy buttons on letter and template blocks
  document.querySelectorAll('[data-copy]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const target = document.getElementById(btn.dataset.copy);
      if (!target) return;
      const text = target.innerText.trim();
      try {
        await navigator.clipboard.writeText(text);
      } catch {
        const ta = document.createElement('textarea');
        ta.value = text; document.body.appendChild(ta); ta.select();
        try { document.execCommand('copy'); } catch {}
        ta.remove();
      }
      const old = btn.textContent;
      btn.textContent = 'Copied';
      btn.classList.add('is-done');
      setTimeout(() => { btn.textContent = old; btn.classList.remove('is-done'); }, 1800);
    });
  });

  // Highlight the table of contents entry for the section in view
  const links = Array.from(document.querySelectorAll('.toc a[href^="#"]'));
  if (links.length && 'IntersectionObserver' in window) {
    const map = new Map(links.map((a) => [a.getAttribute('href').slice(1), a]));
    const io = new IntersectionObserver((entries) => {
      entries.forEach((e) => {
        if (!e.isIntersecting) return;
        links.forEach((a) => a.classList.remove('is-active'));
        map.get(e.target.id)?.classList.add('is-active');
      });
    }, { rootMargin: '-20% 0px -70% 0px' });
    map.forEach((_, id) => { const h = document.getElementById(id); if (h) io.observe(h); });
  }

  // Blog index filters
  const filterBtns = document.querySelectorAll('.filters button');
  filterBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      const f = btn.dataset.filter;
      filterBtns.forEach((b) => b.setAttribute('aria-pressed', String(b === btn)));
      document.querySelectorAll('.post-card').forEach((c) => { c.hidden = f !== 'all' && c.dataset.topic !== f; });
    });
  });

  const y = document.getElementById('year');
  if (y) y.textContent = new Date().getFullYear();
})();
