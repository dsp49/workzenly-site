#!/usr/bin/env python3
"""Build /blog/index.html and /sitemap.xml from the posts and feature pages that exist on disk.

Run after adding or editing posts:  python content-plan/build_index.py
"""
import html, json, os, re, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLAN = json.load(open(os.path.join(ROOT, "content-plan", "plan.json"), encoding="utf-8"))
TODAY = sys.argv[1] if len(sys.argv) > 1 else "2026-10-31"
TOPIC_ORDER = ["Pay and overtime", "Payroll and tax forms", "Minimum wage", "Time off and attendance", "Scheduling", "Time tracking", "Hiring and people", "HR software"]
BATCH_TOPIC = {
    "A_pay_calculators": "Pay and overtime", "B_payroll_forms": "Payroll and tax forms", "C_minimum_wage": "Minimum wage",
    "D_leave_attendance": "Time off and attendance", "E_scheduling": "Scheduling", "F_time_tracking": "Time tracking",
    "G_people_docs": "Hiring and people", "H_software": "HR software",
}


def grab(src, pattern):
    m = re.search(pattern, src, re.S | re.I)
    return html.unescape(re.sub(r"<[^>]+>", "", m.group(1))).strip() if m else ""


posts = []
for p in PLAN["posts"]:
    f = os.path.join(ROOT, "blog", p["slug"], "index.html")
    if not os.path.isfile(f):
        continue
    src = open(f, encoding="utf-8").read()
    posts.append(dict(
        slug=p["slug"], url=p["url"], intent=p["intent"], topic=BATCH_TOPIC[p["batch"]],
        h1=re.sub(r"\s+", " ", grab(src, r"<h1[^>]*>(.*?)</h1>")),
        desc=grab(src, r'<meta name="description" content="(.*?)"'),
    ))
features = [f for f in PLAN["features"] if os.path.isfile(os.path.join(ROOT, "features", f["slug"], "index.html"))]

topics = [t for t in TOPIC_ORDER if any(p["topic"] == t for p in posts)]
slugify = lambda t: re.sub(r"[^a-z]+", "-", t.lower()).strip("-")
esc = html.escape

cards = "\n".join(
    f'      <a class="post-card" href="{p["url"]}" data-topic="{slugify(p["topic"])}">\n'
    f'        <div><span class="post-card__topic">{esc(p["topic"])}</span><b>{esc(p["h1"])}</b></div>\n'
    f'        <p>{esc(p["desc"][:120].rsplit(" ", 1)[0])}…</p>\n'
    f'      </a>'
    for p in sorted(posts, key=lambda p: (TOPIC_ORDER.index(p["topic"]), p["intent"] != "transactional", p["h1"]))
)
filters = '      <button type="button" data-filter="all" aria-pressed="true">All guides</button>\n' + "\n".join(
    f'      <button type="button" data-filter="{slugify(t)}" aria-pressed="false">{esc(t)}</button>' for t in topics)
itemlist = json.dumps({
    "@context": "https://schema.org",
    "@graph": [
        {"@type": "Blog", "name": "Workzenly blog", "url": "https://workzenly.com/blog/",
         "description": "Free HR guides, templates and calculators for US employers and employees."},
        {"@type": "ItemList", "itemListElement": [
            {"@type": "ListItem", "position": i + 1, "url": "https://workzenly.com" + p["url"], "name": p["h1"]}
            for i, p in enumerate(posts)]},
        {"@type": "BreadcrumbList", "itemListElement": [
            {"@type": "ListItem", "position": 1, "name": "Home", "item": "https://workzenly.com/"},
            {"@type": "ListItem", "position": 2, "name": "Blog", "item": "https://workzenly.com/blog/"}]},
    ]}, ensure_ascii=False, indent=1)

page = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>HR Guides, Templates and Calculators · Workzenly Blog</title>
  <meta name="description" content="Free HR guides for US employers: overtime and paycheck calculators, minimum wage by state, W-4 and I-9 help, shift schedules and templates.">
  <meta name="theme-color" content="#0b0b0b">
  <link rel="canonical" href="https://workzenly.com/blog/">
  <link rel="icon" href="/favicon.ico" sizes="32x32">
  <link rel="icon" href="/favicon.svg" type="image/svg+xml">
  <link rel="apple-touch-icon" href="/apple-touch-icon.png">
  <meta property="og:type" content="website">
  <meta property="og:site_name" content="Workzenly">
  <meta property="og:url" content="https://workzenly.com/blog/">
  <meta property="og:title" content="HR Guides, Templates and Calculators · Workzenly Blog">
  <meta property="og:description" content="Free HR guides for US employers: overtime and paycheck calculators, minimum wage by state, W-4 and I-9 help, shift schedules and templates.">
  <meta property="og:image" content="https://workzenly.com/og-image.png">
  <meta name="twitter:card" content="summary_large_image">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Schibsted+Grotesk:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="/styles.css">
  <link rel="stylesheet" href="/blog.css">
  <script type="application/ld+json">
{itemlist}
  </script>
</head>
<body>

<header class="nav">
  <a class="wordmark" href="/" aria-label="Workzenly home">Workzenly</a>
  <nav class="nav__links" aria-label="Main">
    <a class="link" href="/features/time-clock/">Time clock</a>
    <a class="link" href="/features/scheduling/">Scheduling</a>
    <a class="link" href="/features/payroll/">Payroll</a>
    <a class="link" href="/features/pto/">PTO</a>
    <a class="link" href="/blog/" aria-current="page">Blog</a>
  </nav>
  <a class="button button--small" href="/#start">Book a demo</a>
</header>

<main>
  <header class="index-head">
    <h1>HR guides, templates and calculators</h1>
    <p>Free, practical help for US employers and employees: work out overtime and hourly pay, check your state’s minimum wage, pick a shift schedule or copy a template.</p>
  </header>
  <div class="filters" role="group" aria-label="Filter by topic">
{filters}
  </div>
  <section class="post-grid" aria-label="All guides">
{cards}
  </section>

  <section class="close-cta" aria-labelledby="close-title">
    <h2 id="close-title">Run time, schedules and pay in one place</h2>
    <div class="close-cta__side">
      <p>Workzenly handles time clocks, schedules, PTO and payroll for employers, and gives every employee one app for all of it.</p>
      <a class="button" href="/#start">Book a demo</a>
    </div>
  </section>
</main>

<footer class="simple-foot">
  <p>© <span id="year">2026</span> Workzenly</p>
  <nav aria-label="Footer">
    <a href="/features/time-clock/">Time clock</a>
    <a href="/features/scheduling/">Scheduling</a>
    <a href="/features/payroll/">Payroll</a>
    <a href="/features/pto/">PTO</a>
    <a href="/blog/">Blog</a>
  </nav>
  <p><a href="mailto:hello@workzenly.com">hello@workzenly.com</a></p>
</footer>
<script src="/blog.js" defer></script>
</body>
</html>
"""
os.makedirs(os.path.join(ROOT, "blog"), exist_ok=True)
open(os.path.join(ROOT, "blog", "index.html"), "w", encoding="utf-8").write(page)

urls = [("/", "1.0"), ("/blog/", "0.8")] + [(f["url"], "0.9") for f in features] + [(p["url"], "0.7") for p in posts]
sm = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
for u, pr in urls:
    sm.append(f"  <url>\n    <loc>https://workzenly.com{u}</loc>\n    <lastmod>{TODAY}</lastmod>\n    <priority>{pr}</priority>\n  </url>")
sm.append("</urlset>")
open(os.path.join(ROOT, "sitemap.xml"), "w", encoding="utf-8").write("\n".join(sm) + "\n")
print(f"blog index: {len(posts)} posts, {len(topics)} topics; sitemap: {len(urls)} URLs")
