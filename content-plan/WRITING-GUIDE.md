# Workzenly content writing guide (US audience)

Workzenly is workforce software for **US small and mid-sized businesses**, especially hourly and shift-based
teams: restaurants, retail, healthcare, construction, field services, warehouses and offices. It covers the
employee time clock (phone or shared tablet, GPS and geofencing), shift scheduling with swaps, PTO and leave
tracking, payroll with pay stubs in the app, onboarding with e-signed forms, documents and performance reviews,
and gives every employee one app for all of it.

Every reader is in the United States. Write for a US owner, manager or HR lead, and for US employees checking
their pay and rights. Never mention India, rupees, PF, ESI, CTC, gratuity or other non-US concepts.

Book-a-demo is the only conversion: `/#start`. Feature pages: `/features/time-clock/`, `/features/scheduling/`,
`/features/payroll/`, `/features/pto/`.

## Files you work from

- `content-plan/plan.json`: every post (slug, primary keyword, intent, secondary keywords, tool flag, related feature) and every feature page.
- `content-plan/templates/blog-post.html` and `content-plan/templates/feature-page.html`: copy these and keep every structural class name.
- `blog/overtime-calculator/index.html`: the finished reference post. Match its depth, tone, structure, CTA placement and calculator code style.
- `content-plan/seo_check.py`: run `python content-plan/seo_check.py blog/<slug>/index.html` until it prints PASS.
- Output: `blog/<slug>/index.html` or `features/<slug>/index.html`. Edit no other file.

## Match the search intent

- **Informational** (“what is”, “how to”, rules, a state’s minimum wage): answer in the Quick answer box and the first paragraph, then explain with a worked example, a table where it helps, and what employers should do. ≥ 1300 words.
- **Transactional** (calculator, template, checklist): give the thing first. A calculator sits right after the intro paragraph and works on load with sensible defaults. A template post gives 2–4 complete, copyable templates in `.template` blocks with `data-copy` buttons, covering the common variants. A checklist gives the full checklist. Then explain how to use it. ≥ 1000 words.
- **Commercial** (best X, X software, alternatives): help a buyer choose. What the software must do, a buying checklist, a comparison table of criteria, honest trade-offs, and where Workzenly fits. ≥ 1500 words.
  - When you name real products (When I Work, Homebase, Deputy, 7shifts, Sling, Connecteam, Gusto, ADP, Paychex, BambooHR, Rippling, QuickBooks Payroll, etc.), describe only widely known positioning in neutral words. No prices, ratings, user counts, “best/worst” verdicts or feature claims you cannot verify. Tell readers to check current pricing and features on each vendor’s site.
  - Never claim Workzenly has a price, free plan, certification, integration, customer, review or statistic. Describe only what it does (listed above).

## Accuracy: verify every rule and number

US employment rules vary by federal, state and city law and change every year. Use the WebSearch and WebFetch tools (load them with ToolSearch) to confirm current figures from official sources before you print them: dol.gov, irs.gov, uscis.gov, eeoc.gov, and the state labor department (for example dir.ca.gov, floridajobs.org, twc.texas.gov, labor.ny.gov, ohio.gov, michigan.gov).

- Minimum wage pages: give the current rate in effect today (31 October 2026), the effective date, the next scheduled change if officially announced, the tipped minimum and tip credit, the minor/training wage if any, and local/city rates where they exist. Cite the state agency by name in the text.
- FLSA basics: non-exempt employees get overtime at 1.5× the regular rate for hours over 40 in a workweek; the federal minimum wage is $7.25; exempt status needs both the salary level and duties tests. Check the current federal salary threshold status and state thresholds (California, New York, Washington, Colorado and others are higher).
- Forms: W-4, I-9 and E-Verify details must match the current versions on irs.gov and uscis.gov (form edition dates, document lists, deadlines).
- Calculators: check every worked example against the calculator output by tracing your JavaScript with the defaults. Payroll-tax estimates are hard; if you include taxes, label them as estimates and say which year’s IRS figures you used.
- No invented statistics, surveys, quotes, customer stories or “according to” sources. If you cannot confirm a figure, describe the principle and tell readers to check their state labor department.
- Money in USD: $1,234.56. Dates: October 31, 2026. US spelling (labor, organization, color).

## Structure and SEO

- `<title>` 35–60 characters with the primary keyword, written to earn the click (add 2026 where freshness matters: minimum wage, forms, calculators). Meta description 120–158 characters with the keyword and a concrete promise.
- One `<h1>` containing the primary keyword. Keyword in the first 100 words. Secondary keywords in h2s and body, naturally; density under 3%.
- ≥ 5 `<h2>` with ids, all in the table of contents; the last h2 is `id="faq"` “Frequently asked questions”.
- Quick answer box with 3 takeaways at the top.
- **CTAs:** one `.cta-inline` in the middle of the article (between 30% and 65% of the words; the checker measures it) and the `.cta-inline--quiet` Book-a-demo block after the FAQ. The middle CTA links to the related feature page from the plan with a specific title and one sentence on the problem Workzenly removes; the button names the action (“See the time clock”, “See scheduling”, “See payroll”, “See PTO tracking”). Optionally a header aside CTA as in the reference.
- ≥ 4 contextual internal links in the body to other plan URLs, with descriptive anchors. Link only to URLs in plan.json. Three related posts at the bottom, same topic.
- FAQ: 4–6 real questions, direct answers of 40–70 words, mirrored word for word in the FAQPage JSON-LD.
- JSON-LD as in the template; `headline` = the h1 text.

## Voice

- Plain, specific, friendly. Short paragraphs. Busy US small-business owners and managers, and hourly employees.
- Specific nouns and real numbers. No filler intros, no “In conclusion”.
- Avoid: seamless, unlock, empower, leverage, robust, cutting-edge, game-changer, delve, elevate, revolutionize, all-in-one, “look no further”. Few or no em dashes.
- Sentence-case headings.

## Listicles (/tools/ and /templates/)

Listicles are the pages we most want to rank. They must be the most useful page on the first results page
for their keyword: complete, current, honest and quick to scan.

### Software roundups: /tools/<slug>/ (template: content-plan/templates/tool-listicle.html)

1. **Research the field first.** Use WebSearch to find the products that really compete for this query in the
   US today (category pages on G2 or Capterra are fine for finding candidates, never for copying ratings or
   quotes). Pick 7 to 10 that fit the searcher, all currently operating, with a US website.
2. **Verify every product on its own site today** with WebFetch: home page, features, pricing page, app store
   listing if relevant. Record: what it is, who it is built for, the 2 to 4 features that matter for this query,
   free plan or trial (and its limits), the lowest published price with its unit (per user/month, per
   location/month, base fee + per person), platforms (web, iOS, Android, kiosk). If no price is published,
   write "Quote". Never guess; if a fact cannot be confirmed, leave it out.
3. **Capture assets for every product** (shared across listicles, so reuse folders that already exist):
   `python content-plan/capture_tool.py <tool-slug> <homepage-url> --name="<Name>"`. Then open
   `assets/tools/<tool-slug>/home.jpg` and `logo.png` with the Read tool and look at them.
   - If the screenshot is a captcha, an error, or mostly a cookie wall: re-run with `--force` once; if still bad,
     try the product’s features page URL instead of the home page.
   - If the logo is a lettermark, blurry, or the wrong mark: re-run with `--logo-url=<official PNG/JPG>` (press
     kit, brand page, or a Wikimedia Commons PNG thumbnail), or crop the brand mark out of the screenshot with
     `--crop-logo=x,y,w,h` (1280×720 coordinates; check the result).
4. **Order** the list by best fit for the searcher, and say in each card who it is best for. Each “Top pick” covers
   a different buyer (best overall, best free, best for restaurants, best for payroll, and so on).
5. **Workzenly appears once**, as an honest entry: `tool-card tool-card--ours`, the “Our product” tag, logo
   `/apple-touch-icon.png`, screenshot `/og-image.png` (1200×630). Describe only what it does (see the top of this
   guide). Its price and trial fields say “Book a demo”. On “X alternatives” pages, never list X itself as an
   alternative; give X a short fair summary in the intro instead.
6. **Each tool card** has 180–260 words: what it is and who it suits, what stands out (specific features, not
   adjectives), 3–4 strengths and 2–3 drawbacks that are concrete and checkable (“Payroll is a paid add-on”,
   “No Android app”, “Free plan limited to one location”). No star ratings, scores, user counts, awards or review
   quotes. No disparaging language; drawbacks are facts.
7. **Sections in order:** disclosure, intro (keyword in the first 100 words; who this list is for), top picks,
   comparison table, the tool cards (mid-article CTA after the 3rd or 4th card), how we chose, how to choose
   (a buyer’s checklist specific to the category, with US specifics like FLSA overtime, state breaks, predictive
   scheduling laws where relevant), FAQ, closing CTA. Related links at the bottom go to other /tools/ and
   /templates/ pages.
8. **Alt text**: logos “<Name> logo”; screenshots “<Name> home page showing <what is visible>”.

### Free-template roundups: /templates/<slug>/ (template: content-plan/templates/template-listicle.html)

1. Offer 5 to 10 genuinely different templates (layout, use case, industry, period), not one template five times.
2. Build the real files with `content-plan/tplkit.py` from a script saved at
   `content-plan/template_builds/<slug>.py` (so they can be rebuilt). Files go to `assets/templates/<slug>/` with
   clear names (`weekly-employee-schedule-restaurant.xlsx`). Sheets: .xlsx + printable .pdf. Letters and forms:
   .docx + .pdf + a “Copy text” button. Use working formulas for totals and hours where the template needs them.
3. Run the build script, then open at least one .pdf and one .xlsx/.docx to confirm they were created.
4. The on-page preview shows the template filled with realistic US sample data, so a visitor sees exactly what
   they get before downloading. Wide previews scroll inside their box on phones.
5. Each template card: name, best for, preview, download buttons, a one-line description of what is inside, and
   two or three sentences on when and how to use it. Letters and HR forms add the key legal notes (for example,
   at-will language, final paycheck timing varies by state, keep records for the required period) as principles,
   verified on dol.gov or the state labor site, and tell readers to check their state.
