#!/usr/bin/env python3
"""Build the hub pages (/blog/, /tools/, /templates/) and /sitemap.xml from the pages that exist on disk.

Run after adding or editing pages:  python content-plan/build_index.py [YYYY-MM-DD]
"""
import html, json, os, re, sys, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLAN = json.load(open(os.path.join(ROOT, "content-plan", "plan.json"), encoding="utf-8"))
TODAY = sys.argv[1] if len(sys.argv) > 1 else datetime.date.today().isoformat()
esc = html.escape
slugify = lambda t: re.sub(r"[^a-z]+", "-", t.lower()).strip("-")

BLOG_TOPICS = {
    "A_pay_calculators": "Pay and overtime", "B_payroll_forms": "Payroll and tax forms", "C_minimum_wage": "Minimum wage",
    "D_leave_attendance": "Time off and attendance", "E_scheduling": "Scheduling", "F_time_tracking": "Time tracking",
    "G_people_docs": "Hiring and people", "H_software": "HR software",
}
TOOL_KINDS = {"best": "Best software", "free": "Free tools", "alternatives": "Alternatives"}
TEMPLATE_TOPICS = {"scheduling": "Schedules", "time-clock": "Timesheets and attendance", "pto": "Time off and policies",
                   "payroll": "Hiring, pay and letters"}

SECTIONS = {
    "blog": dict(
        title="HR Guides and Calculators for US Employers · Workzenly",
        desc="Free HR guides for US employers: overtime and paycheck calculators, minimum wage by state, W-4 and I-9 help, shift schedules and leave rules.",
        h1="HR guides and calculators",
        lead="Practical answers for US employers and employees: work out overtime and hourly pay, check your state’s minimum wage, and get leave, scheduling and payroll rules right.",
        crumb="Guides", all_label="All guides"),
    "tools": dict(
        title="Best HR, Scheduling and Payroll Software Compared · Workzenly",
        desc="Researched comparisons of time clock apps, employee scheduling apps, payroll and HR software for US small businesses, with prices, screenshots and free options.",
        h1="Best tools for hourly teams",
        lead="Side-by-side comparisons of time clock, scheduling, PTO, payroll and HR software for US small businesses. Every product checked on its own website, with screenshots and current prices.",
        crumb="Best tools", all_label="All comparisons"),
    "templates": dict(
        title="Free HR Templates: Schedules, Timesheets, Forms · Workzenly",
        desc="Free HR templates for US employers: weekly schedules, timesheets, PTO trackers, offer and termination letters, write-up forms. Excel, Word and PDF, no sign-up.",
        h1="Free HR templates",
        lead="Download ready-to-use schedules, timesheets, trackers, letters and forms in Excel, Word and PDF. No sign-up, free to edit.",
        crumb="Free templates", all_label="All templates"),
}


def grab(src, pattern):
    m = re.search(pattern, src, re.S | re.I)
    return html.unescape(re.sub(r"<[^>]+>", "", m.group(1))).strip() if m else ""


def collect(section):
    out = []
    for p in PLAN["posts"]:
        if p.get("section", "blog") != section:
            continue
        f = os.path.join(ROOT, section, p["slug"], "index.html")
        if not os.path.isfile(f):
            continue
        src = open(f, encoding="utf-8").read()
        if section == "blog":
            topic = BLOG_TOPICS.get(p.get("batch"), "Guides")
        elif section == "tools":
            topic = TOOL_KINDS.get(p.get("list_kind"), "Best software")
        else:
            topic = TEMPLATE_TOPICS.get(p.get("feature"), "Templates")
        logos = re.findall(r'<img class="tool-card__logo" src="([^"]+)"', src)[:5] if section == "tools" else []
        out.append(dict(slug=p["slug"], url=p["url"], intent=p["intent"], topic=topic, volume=p.get("volume", 0),
                        h1=re.sub(r"\s+", " ", grab(src, r"<h1[^>]*>(.*?)</h1>")),
                        desc=grab(src, r'<meta name="description" content="(.*?)"'), logos=logos))
    return out


def nav(active):
    label = {"blog": "Resources", "tools": "Resources", "templates": "Resources"}[active]
    return f"""<header class="nav">
  <a class="wordmark" href="/" aria-label="Workzenly home">Workzenly</a>
  <nav class="nav__links" aria-label="Main">
    <a class="link" href="/features/time-clock/">Time clock</a>
    <a class="link" href="/features/scheduling/">Scheduling</a>
    <a class="link" href="/features/payroll/">Payroll</a>
    <a class="link" href="/features/pto/">PTO</a>
    <a class="link" href="/blog/" aria-current="page">{label}</a>
  </nav>
  <a class="button button--small" href="/#start">Book a demo</a>
</header>"""


FOOT = """<footer class="simple-foot">
  <p>© <span id="year">2026</span> Workzenly</p>
  <nav aria-label="Footer">
    <a href="/features/time-clock/">Time clock</a>
    <a href="/features/scheduling/">Scheduling</a>
    <a href="/features/payroll/">Payroll</a>
    <a href="/features/pto/">PTO</a>
    <a href="/tools/">Best tools</a>
    <a href="/templates/">Free templates</a>
    <a href="/blog/">Guides</a>
  </nav>
  <p><a href="mailto:info@workzenly.com">info@workzenly.com</a></p>
</footer>"""


def hub(section, items, counts):
    cfg = SECTIONS[section]
    topics = []
    for it in items:
        if it["topic"] not in topics:
            topics.append(it["topic"])
    order = {t: i for i, t in enumerate(topics)}
    items = sorted(items, key=lambda p: (order[p["topic"]], -p["volume"]))
    cards = []
    for p in items:
        logos = ("<span class=\"post-card__logos\">" + "".join(f'<img src="{esc(l)}" width="28" height="28" alt="" loading="lazy">' for l in p["logos"]) + "</span>") if p["logos"] else ""
        cards.append(
            f'      <a class="post-card" href="{p["url"]}" data-topic="{slugify(p["topic"])}">\n'
            f'        <div><span class="post-card__topic">{esc(p["topic"])}</span><b>{esc(p["h1"])}</b></div>\n'
            f'        {logos}<p>{esc(p["desc"][:120].rsplit(" ", 1)[0])}…</p>\n'
            f'      </a>')
    filters = f'      <button type="button" data-filter="all" aria-pressed="true">{cfg["all_label"]}</button>\n' + "\n".join(
        f'      <button type="button" data-filter="{slugify(t)}" aria-pressed="false">{esc(t)}</button>' for t in topics)
    current = ' aria-current="page"'
    tabs = "\n".join(
        f'    <a href="/{s}/"{current if s == section else ""}>{SECTIONS[s]["crumb"]} <span class="tpl-meta" style="margin:0 0 0 6px">{counts[s]}</span></a>'
        for s in ("tools", "templates", "blog") if counts[s])
    url = f"https://workzenly.com/{section}/"
    ld = json.dumps({"@context": "https://schema.org", "@graph": [
        {"@type": "CollectionPage", "name": cfg["h1"], "url": url, "description": cfg["desc"]},
        {"@type": "ItemList", "itemListElement": [
            {"@type": "ListItem", "position": i + 1, "url": "https://workzenly.com" + p["url"], "name": p["h1"]}
            for i, p in enumerate(items)]},
        {"@type": "BreadcrumbList", "itemListElement": [
            {"@type": "ListItem", "position": 1, "name": "Home", "item": "https://workzenly.com/"},
            {"@type": "ListItem", "position": 2, "name": cfg["crumb"], "item": url}]},
    ]}, ensure_ascii=False, indent=1)
    page = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(cfg["title"])}</title>
  <meta name="description" content="{esc(cfg["desc"])}">
  <meta name="theme-color" content="#0b0b0b">
  <link rel="canonical" href="{url}">
  <link rel="icon" href="/favicon.ico" sizes="32x32">
  <link rel="icon" href="/favicon.svg" type="image/svg+xml">
  <link rel="apple-touch-icon" href="/apple-touch-icon.png">
  <meta property="og:type" content="website">
  <meta property="og:site_name" content="Workzenly">
  <meta property="og:url" content="{url}">
  <meta property="og:title" content="{esc(cfg["title"])}">
  <meta property="og:description" content="{esc(cfg["desc"])}">
  <meta property="og:image" content="https://workzenly.com/og-image.png">
  <meta name="twitter:card" content="summary_large_image">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Schibsted+Grotesk:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="/styles.css">
  <link rel="stylesheet" href="/blog.css">
  <script type="application/ld+json">
{ld}
  </script>
</head>
<body>

{nav(section)}

<main>
  <header class="index-head">
    <h1>{esc(cfg["h1"])}</h1>
    <p>{esc(cfg["lead"])}</p>
  </header>
  <nav class="hub-tabs" aria-label="Resource sections">
{tabs}
  </nav>
  <div class="filters" role="group" aria-label="Filter by topic">
{filters}
  </div>
  <section class="post-grid" aria-label="{esc(cfg["all_label"])}">
{chr(10).join(cards)}
  </section>

  <section class="close-cta" aria-labelledby="close-title">
    <h2 id="close-title">Run time, schedules and pay in one place</h2>
    <div class="close-cta__side">
      <p>Workzenly handles time clocks, schedules, PTO and payroll for employers, and gives every employee one app for all of it.</p>
      <a class="button" href="/#start">Book a demo</a>
    </div>
  </section>
</main>

{FOOT}
<script src="/blog.js" defer></script>
</body>
</html>
"""
    os.makedirs(os.path.join(ROOT, section), exist_ok=True)
    open(os.path.join(ROOT, section, "index.html"), "w", encoding="utf-8").write(page)


data = {s: collect(s) for s in ("blog", "tools", "templates")}
counts = {s: len(v) for s, v in data.items()}
for s, items in data.items():
    if items or s == "blog":
        hub(s, items, counts)
features = [f for f in PLAN["features"] if os.path.isfile(os.path.join(ROOT, "features", f["slug"], "index.html"))]

urls = [("/", "1.0")] + [(f["url"], "0.9") for f in features]
for s in ("tools", "templates", "blog"):
    if data[s]:
        urls.append((f"/{s}/", "0.8"))
        urls += [(p["url"], "0.7") for p in data[s]]
sm = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
for u, pr in urls:
    sm.append(f"  <url>\n    <loc>https://workzenly.com{u}</loc>\n    <lastmod>{TODAY}</lastmod>\n    <priority>{pr}</priority>\n  </url>")
sm.append("</urlset>")
open(os.path.join(ROOT, "sitemap.xml"), "w", encoding="utf-8").write("\n".join(sm) + "\n")
print("hubs:", counts, "| features:", len(features), "| sitemap:", len(urls), "URLs")
