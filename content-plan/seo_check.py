#!/usr/bin/env python3
"""Workzenly on-page SEO checker.

Usage:
  python content-plan/seo_check.py blog/gratuity-calculator/index.html [more files...]
  python content-plan/seo_check.py --all          # every page in /blog and /features
  python content-plan/seo_check.py --all --quiet  # only failures and the summary

Exit code 0 when every file passes (warnings allowed), 1 otherwise.
"""
import json, os, re, sys
from html.parser import HTMLParser
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLAN = json.load(open(os.path.join(ROOT, "content-plan", "plan.json"), encoding="utf-8"))
POSTS = {p["slug"]: p for p in PLAN["posts"]}
FEATS = {f["slug"]: f for f in PLAN["features"]}
VALID_PATHS = {"/", "/blog/", "/#start", "/#features"} | {p["url"] for p in PLAN["posts"]} | {f["url"] for f in PLAN["features"]}
MIN_WORDS = {"informational": 1300, "transactional": 1000, "commercial": 1500}
BANNED = ["seamless", "seamlessly", "unlock", "unleash", "empower", "leverage", "game-changer", "game changer",
          "cutting-edge", "fast-paced", "delve", "supercharge", "revolutionize", "robust", "elevate your",
          "in today's", "look no further", "all-in-one"]
VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}


class Node:
    __slots__ = ("tag", "attrs", "children", "parent", "text")

    def __init__(self, tag, attrs, parent):
        self.tag, self.attrs, self.parent, self.children, self.text = tag, dict(attrs), parent, [], []

    def classes(self):
        return (self.attrs.get("class") or "").split()

    def iter(self):
        yield self
        for c in self.children:
            if isinstance(c, Node):
                yield from c.iter()

    def find_all(self, tag=None, cls=None, pred=None):
        for n in self.iter():
            if tag and n.tag != tag:
                continue
            if cls and cls not in n.classes():
                continue
            if pred and not pred(n):
                continue
            yield n

    def find(self, tag=None, cls=None, pred=None):
        return next(self.find_all(tag, cls, pred), None)

    def get_text(self, skip=("script", "style")):
        out = []
        for c in self.children:
            if isinstance(c, Node):
                if c.tag in skip:
                    continue
                out.append(c.get_text(skip))
            else:
                out.append(c)
        return " ".join(out)

    def ancestors(self):
        p = self.parent
        while p is not None:
            yield p
            p = p.parent


class Builder(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("#root", [], None)
        self.cur = self.root
        self.scripts = []

    def handle_starttag(self, tag, attrs):
        n = Node(tag, attrs, self.cur)
        self.cur.children.append(n)
        if tag not in VOID:
            self.cur = n

    def handle_startendtag(self, tag, attrs):
        self.cur.children.append(Node(tag, attrs, self.cur))

    def handle_endtag(self, tag):
        n = self.cur
        while n is not None and n.tag != tag:
            n = n.parent
        if n is not None and n.parent is not None:
            self.cur = n.parent

    def handle_data(self, data):
        self.cur.children.append(data)
        if self.cur.tag == "script":
            self.scripts.append((dict(self.cur.attrs), data))


def norm(s):
    s = s.lower().replace("’", "'").replace("-", " ").replace("/", " ")
    return re.sub(r"\s+", " ", re.sub(r"[^\w\s']", " ", s)).strip()


def has_kw(text, kw):
    t, k = norm(text), norm(kw)
    if k in t:
        return True
    # tolerate light reordering: every keyword token present
    return all(re.search(r"\b" + re.escape(tok), t) for tok in k.split())


def words(text):
    return re.findall(r"[A-Za-z0-9₹%][\w’'%₹.,/-]*", text)


def check(path, quiet=False):
    rel = os.path.relpath(path, ROOT).replace("\\", "/")
    m = re.match(r"(blog|features)/([^/]+)/index\.html$", rel)
    fails, warns = [], []
    if not m:
        return rel, ["path must be blog/<slug>/index.html or features/<slug>/index.html"], [], {}
    kind, slug = m.groups()
    plan = (POSTS if kind == "blog" else FEATS).get(slug)
    if not plan:
        return rel, [f"slug '{slug}' is not in content-plan/plan.json"], [], {}
    kw, intent = plan["keyword"], plan["intent"]
    html = open(path, encoding="utf-8").read()
    b = Builder()
    b.feed(html)
    root = b.root
    info = {}

    for bad in ("{{", "}}", "lorem ipsum", "TODO", "PLACEHOLDER", "Acme"):
        if bad.lower() in html.lower():
            fails.append(f"leftover placeholder text: {bad}")

    # --- head ---
    title_n = root.find("title")
    title = title_n.get_text().strip() if title_n else ""
    info["title"] = title
    if not (35 <= len(title) <= 60):
        fails.append(f"title is {len(title)} chars (need 35–60): {title!r}")
    if not has_kw(title, kw):
        fails.append(f"title lacks primary keyword '{kw}'")

    md_n = root.find("meta", pred=lambda n: n.attrs.get("name") == "description")
    md = (md_n.attrs.get("content") or "").strip() if md_n else ""
    info["description"] = md
    if not (120 <= len(md) <= 158):
        fails.append(f"meta description is {len(md)} chars (need 120–158)")
    if not has_kw(md, kw):
        fails.append(f"meta description lacks primary keyword '{kw}'")

    expected = f"https://workzenly.com/{kind}/{slug}/"
    can = root.find("link", pred=lambda n: n.attrs.get("rel") == "canonical")
    if not can or can.attrs.get("href") != expected:
        fails.append(f"canonical must be {expected}")
    ogu = root.find("meta", pred=lambda n: n.attrs.get("property") == "og:url")
    if not ogu or ogu.attrs.get("content") != expected:
        fails.append("og:url must equal the canonical URL")
    for prop in ("og:title", "og:description", "og:image"):
        if not root.find("meta", pred=lambda n, p=prop: n.attrs.get("property") == p):
            fails.append(f"missing {prop}")
    if root.find("meta", pred=lambda n: n.attrs.get("name") == "robots" and "noindex" in (n.attrs.get("content") or "")):
        fails.append("page is noindex")

    body = root.find("body")
    if not body or norm(body.attrs.get("data-keyword", "")) != norm(kw):
        fails.append(f"<body data-keyword> must be '{kw}'")
    if not body or body.attrs.get("data-intent") != intent:
        fails.append(f"<body data-intent> must be '{intent}'")

    # --- headings ---
    h1s = list(root.find_all("h1"))
    if len(h1s) != 1:
        fails.append(f"{len(h1s)} <h1> elements (need exactly 1)")
    elif not has_kw(h1s[0].get_text(), kw):
        fails.append(f"<h1> lacks primary keyword '{kw}'")
    else:
        h1t = re.sub(r"\s+", " ", h1s[0].get_text()).strip()
        info["h1"] = h1t
        if len(h1t) > 70:
            warns.append(f"<h1> is long ({len(h1t)} chars)")

    main = root.find("main")
    content = root.find("article", cls="prose") if kind == "blog" else main
    if content is None:
        fails.append("no article.prose (blog) or <main> (feature) found")
        return rel, fails, warns, info
    h2s = [h for h in content.find_all("h2")]
    if kind == "feature" or kind == "features":
        h2s = [h for h in main.find_all("h2")]
    if len(h2s) < 5:
        fails.append(f"only {len(h2s)} <h2> headings (need ≥ 5)")
    if kind == "blog":
        ids = [h.attrs.get("id") for h in h2s]
        if not all(ids):
            fails.append("every article <h2> needs an id")
        toc = root.find("nav", cls="toc")
        toc_ids = [a.attrs.get("href", "")[1:] for a in toc.find_all("a")] if toc else []
        missing = [i for i in ids if i and i not in toc_ids]
        if missing:
            fails.append(f"h2 ids missing from table of contents: {missing}")
        dead = [i for i in toc_ids if i not in ids]
        if dead:
            fails.append(f"table of contents links to missing ids: {dead}")
        dup = [i for i, c in Counter(ids).items() if c > 1]
        if dup:
            fails.append(f"duplicate h2 ids: {dup}")

    # --- words, keyword placement ---
    text = content.get_text()
    if kind == "blog":
        # body text excluding the CTA blocks and the byline
        skip_nodes = set(id(n) for n in content.find_all(cls="cta-inline")) | set(id(n) for n in content.find_all(cls="byline"))

        def txt(n):
            out = []
            for c in n.children:
                if isinstance(c, Node):
                    if c.tag in ("script", "style") or id(c) in skip_nodes:
                        continue
                    out.append(txt(c))
                else:
                    out.append(c)
            return " ".join(out)
        text = txt(content)
    w = words(text)
    info["words"] = len(w)
    need = MIN_WORDS[intent] if kind == "blog" else 900
    if len(w) < need:
        fails.append(f"{len(w)} words (need ≥ {need} for {intent})")
    first100 = " ".join(w[:100])
    if not has_kw(first100, kw):
        fails.append(f"primary keyword '{kw}' not in the first 100 words")
    kcount = len(re.findall(re.escape(norm(kw)), norm(text)))
    density = kcount * len(norm(kw).split()) / max(1, len(w)) * 100
    info["kw_count"] = kcount
    if kcount < 3:
        warns.append(f"primary keyword phrase appears only {kcount}× in the body")
    if density > 3.0:
        fails.append(f"keyword stuffing: density {density:.1f}% (keep under 3%)")
    sec_hits = [s for s in plan.get("secondary", []) if has_kw(text, s)]
    if len(sec_hits) < min(2, len(plan.get("secondary", []))):
        warns.append(f"few secondary keywords used ({len(sec_hits)}/{len(plan.get('secondary', []))})")

    # --- CTAs ---
    ctas = list((content if kind == "blog" else main).find_all(cls="cta-inline"))
    need_cta = 2 if kind == "blog" else 1
    if len(ctas) < need_cta:
        fails.append(f"{len(ctas)} .cta-inline blocks (need ≥ {need_cta})")
    for c in ctas:
        a = c.find("a")
        if not a or not a.attrs.get("href"):
            fails.append("a CTA block has no link")
    if kind == "blog" and ctas:
        # position of the first CTA, measured in body words before it
        marker = "⁣CTA⁣"
        first = ctas[0]
        first.children.insert(0, marker)
        before = content.get_text().split(marker)[0]
        first.children.pop(0)
        pos = len(words(before)) / max(1, len(words(content.get_text())))
        info["cta_pos"] = round(pos, 2)
        if not (0.28 <= pos <= 0.67):
            fails.append(f"first CTA sits at {pos:.0%} of the article (place it between 30% and 65%)")

    # --- links ---
    scope = content if kind == "blog" else main
    in_links = set()
    for a in scope.find_all("a"):
        href = a.attrs.get("href", "")
        if any(x.tag in ("nav", "footer") or "related" in x.classes() for x in a.ancestors()):
            continue
        if href.startswith(("/blog/", "/features/")) and href.rstrip("#") != f"/{kind}/{slug}/":
            in_links.add(href.split("#")[0])
    info["internal_links"] = len(in_links)
    if len(in_links) < 4:
        fails.append(f"{len(in_links)} contextual internal links (need ≥ 4 to other /blog/ or /features/ pages)")
    for a in root.find_all("a"):
        href = a.attrs.get("href", "")
        if href.startswith("/") and not href.startswith("//"):
            base = href.split("?")[0]
            p = base.split("#")[0] or "/"
            if p not in VALID_PATHS and not os.path.exists(os.path.join(ROOT, p.lstrip("/"))):
                fails.append(f"broken internal link: {href}")
        elif href.startswith("http") and "workzenly.com" not in href:
            if "noopener" not in (a.attrs.get("rel") or "") and a.attrs.get("target") == "_blank":
                warns.append(f"external link opens new tab without rel=noopener: {href}")
        if not a.get_text().strip() and not a.attrs.get("aria-label"):
            fails.append(f"empty link text: {href}")
    for img in root.find_all("img"):
        if img.attrs.get("alt") is None:
            fails.append(f"image without alt: {img.attrs.get('src')}")

    # --- FAQ + JSON-LD ---
    faq = (content if kind == "blog" else main).find(cls="faq-list")
    qs = [re.sub(r"\s+", " ", d.find("summary").get_text()).strip() for d in faq.find_all("details") if d.find("summary")] if faq else []
    lo, hi = (4, 7) if kind == "blog" else (5, 8)
    if not (lo <= len(qs) <= hi):
        fails.append(f"{len(qs)} FAQ questions (need {lo}–{hi})")
    types, ld_qs = set(), []
    lds = [s for a, s in b.scripts if a.get("type") == "application/ld+json"]
    if not lds:
        fails.append("no JSON-LD")
    for s in lds:
        try:
            data = json.loads(s)
        except Exception as e:
            fails.append(f"JSON-LD does not parse: {e}")
            continue
        items = data.get("@graph", [data]) if isinstance(data, dict) else data
        for it in items:
            t = it.get("@type")
            types.add(t)
            if t == "FAQPage":
                ld_qs = [q.get("name", "").strip() for q in it.get("mainEntity", [])]
                for q in it.get("mainEntity", []):
                    if not q.get("acceptedAnswer", {}).get("text", "").strip():
                        fails.append("FAQPage answer is empty")
            if t == "BlogPosting":
                if it.get("mainEntityOfPage") != expected:
                    fails.append("BlogPosting.mainEntityOfPage must equal the canonical URL")
                if info.get("h1") and norm(it.get("headline", "")) != norm(info["h1"]):
                    warns.append("BlogPosting.headline differs from the <h1>")
    need_types = {"BlogPosting", "BreadcrumbList", "FAQPage"} if kind == "blog" else {"WebPage", "BreadcrumbList", "FAQPage"}
    if not need_types <= types:
        fails.append(f"JSON-LD missing types: {sorted(need_types - types)}")
    if qs and [norm(q) for q in qs] != [norm(q) for q in ld_qs]:
        fails.append("FAQ questions on the page and in FAQPage JSON-LD differ")

    # --- tools ---
    if kind == "blog" and plan.get("tool"):
        if not content.find(cls="tool"):
            fails.append("calculator post needs a working .tool widget")
        if not any(not a.get("src") and a.get("type") != "application/ld+json" and s.strip() for a, s in b.scripts):
            fails.append("calculator post needs an inline <script> that powers the tool")
    if kind == "features" and not root.find(pred=lambda n: "data-ui" in n.attrs):
        fails.append("feature page hero needs a product fragment with data-ui")

    # --- style / copy hygiene ---
    low = norm(text)
    hits = [w_ for w_ in BANNED if norm(w_) in low]
    if hits:
        warns.append(f"generic marketing words: {hits}")
    dashes = text.count("—")
    if dashes > 3:
        warns.append(f"{dashes} em dashes; prefer full stops")
    for s in ("/styles.css", "/blog.css"):
        if s not in html:
            fails.append(f"stylesheet {s} not linked")
    if '"/blog.js"' not in html:
        warns.append("/blog.js not loaded")
    return rel, fails, warns, info


def main(argv):
    quiet = "--quiet" in argv
    args = [a for a in argv if not a.startswith("--")]
    if "--all" in argv:
        args = []
        for kind in ("blog", "features"):
            d = os.path.join(ROOT, kind)
            if os.path.isdir(d):
                for s in sorted(os.listdir(d)):
                    p = os.path.join(d, s, "index.html")
                    if os.path.isfile(p):
                        args.append(p)
    if not args:
        print(__doc__)
        return 2
    results = [check(os.path.abspath(p)) for p in args]
    titles, descs = defaultdict(list), defaultdict(list)
    for rel, f, w, info in results:
        if info.get("title"):
            titles[info["title"].lower()].append(rel)
        if info.get("description"):
            descs[info["description"].lower()].append(rel)
    for group, label in ((titles, "title"), (descs, "meta description")):
        for k, files in group.items():
            if len(files) > 1:
                for rel, f, w, info in results:
                    if rel in files:
                        f.append(f"duplicate {label} shared with {[x for x in files if x != rel]}")
    bad = 0
    for rel, f, w, info in results:
        status = "PASS" if not f else "FAIL"
        bad += bool(f)
        if quiet and not f:
            continue
        extra = f"  words={info.get('words')} links={info.get('internal_links')} cta@{info.get('cta_pos')}" if info else ""
        print(f"{status}  {rel}{extra}")
        for x in f:
            print(f"   FAIL  {x}")
        for x in w:
            print(f"   warn  {x}")
    print(f"\n{len(results) - bad}/{len(results)} pages pass")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
