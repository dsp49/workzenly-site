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
