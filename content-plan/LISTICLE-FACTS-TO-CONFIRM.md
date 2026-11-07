# Listicles: facts to confirm before publishing

The auditor agents flagged these for a human check. Prices and plan limits change often; confirm on each vendor's pricing page. Legal notes are written as principles; confirm with your state labor department or an employment attorney.


## pilot_templates

- FMLA: covers private employers with 50+ employees in 20+ workweeks; up to 12 workweeks; at least 30 days' notice for foreseeable leave; eligibility notice within 5 business days; employees generally get at least 15 calendar days to return certification (DOL Fact Sheets 28, 28E and 28G). Confirm nothing has changed before publishing.
- Oregon predictive scheduling: BOLI's pages describing the written-employee-request exemption appear to be several years old, and a 2026 BOLI publication mentions recent updates. Re-check the current ORS 653.455 and 653.436 text.
- California paid sick leave (40 hours or 5 days, whichever is more) and the vacation vesting / no use-it-or-lose-it rule come from dir.ca.gov pages. Re-confirm before publishing in case of 2026 updates.
- FLSA record keeping (payroll records 3 years, time cards 2 years) and FMLA records (3 years, medical records in a separate confidential file) come from DOL Fact Sheet 21 and 29 CFR 825.500. These are stable, but worth a final check.
- The partial-day exempt salary deduction statement (full-day absences only) is based on DOL Fact Sheet 17G.
- Notice lead times on the page ('two weeks for a few days, 30 days for a week or more') are presented as common policy practice, not law.
- Five internal links (/blog/bereavement-leave/, /blog/jury-duty-pay/, /blog/holiday-pay/, /blog/attendance-policy/, /blog/fmla-leave/) and several /templates/ and /tools/ links point to planned pages that are not written yet. Make sure they exist before or at launch, or temporarily remove the links. /features/pto/ (the CTA target) and /templates/ (breadcrumb) also do not exist yet.

## pilot_tools

- BLOCKER: /og-image.png (Workzenly's screenshot, required by the guide) shows a payslip in rupees ('₹48,260') plus 'Leave left' and 'payslip' wording. This breaks the US-only rule. Regenerate it in USD with US wording ('PTO left', 'Next pay stub') before publishing. It is a shared file outside this page's scope.
- Buddy Punch: buddypunch.com/pricing returns 403 to automated tools. The $19 monthly base fee + $4.99 per user (Starter), geofencing on Pro, the $2 per user real-time GPS add-on and scheduling as a Starter add-on were confirmed only through docs.buddypunch.com and search snippets. Third-party sites disagree on whether $4.99 is the annual or monthly rate (some say $5.99 monthly). Check the live pricing page in a browser.
- Homebase $30 per location (Essentials) and Connecteam $29 for the first 30 users: both pricing pages have an annual/monthly toggle and do not label which billing period the shown price uses. Connecteam's other figure is $35, which is probably the monthly rate.
- QuickBooks Time $20/month + $10 per user is the list price; a promo currently cuts the base fee 50% for 3 months. Confirm the trial terms for the Time plans specifically.
- Deputy's prices ($5/$6.50/$9) are not labeled monthly or annual on its pricing page (there is a 10% annual discount toggle). The 'up to 31 days, no card' trial comes from the home page.
- When I Work: the pricing page shows time and attendance as a toggle on each plan, but its FAQ says every plan includes it. Confirm the real per-user cost with the time clock turned on.
- Clockify's free plan is listed as 'up to 5 users' in the header and '4 people' in the FAQ. The page says 'a handful of users', which is safe either way.
- Jibble's paid tier prices and the split between free and paid features are not on its pricing page. The page says only 'priced per user'.
- Connecteam's iOS/Android apps and web dashboard are well known, but the pricing page does not list platforms.
- FAQ range 'about $2.50 to $13 per user per month' depends on Buddy Punch's top tier at $12.99 (from search results). Re-check it if Buddy Punch changes its prices.
- Legal notes stated as principles and not re-checked on official sources today: Illinois BIPA written notice and release; the predictive scheduling list (Oregon, NYC, Chicago, Seattle, San Francisco, Los Angeles, Philadelphia); California daily overtime after 8 hours and double time after 12; neutral rounding (29 CFR 785.48). The FLSA rule to keep payroll records 3 years and time cards 2 years was verified by the writer on DOL Fact Sheet 21.

## P1_schedule_templates

- California alternative workweek: the new line (an alternative workweek can raise the daily overtime threshold to up to 10 hours, or 12 in health care, and hours over 12 stay double time) comes from IWC wage-order text in search results on dir.ca.gov. I did not open the current wage order itself. Have someone check it before publishing.
- Seattle Secure Scheduling (confirmed on seattle.gov today): covers retail and food service employers with 500 or more employees worldwide (full-service restaurants also need 40 or more locations), requires schedules 14 days ahead, and pays time-and-a-half for hours worked between closing and opening shifts less than 10 hours apart.
- Oregon predictive scheduling (confirmed on oregon.gov/boli today): 500 or more employees worldwide in retail, hospitality or food service; schedules 14 days ahead; one extra hour of pay for employer additions of more than 30 minutes or for moved shifts; half pay for scheduled hours cut; changes an employee asks for in writing are exempt, unless they fall within 10 hours of the previous shift. The change-log example relies on these rules.
- DOL Fact Sheet 21 (confirmed today): payroll records kept 3 years; time cards and work time schedules kept 2 years.
- Not fetched again in this audit; I relied on the writer's checks: the federal hour limits for 14- and 15-year-olds, the 8-and-80 system (Fact Sheet 54), on-call time (Fact Sheet 22), and the workweek-change rule (29 CFR 778.105).
- The 12-hour rotation definitions (Pitman, DuPont, 4-on-4-off, 7-on-7-off) come from industry sources, not an official definition. They match each other and the math checks out.
- Re-running the weekly and employee build scripts rebuilt every file in those two asset folders. They are blank templates and only the two PDFs I fixed changed in a meaningful way; you could spot-check one .xlsx from each in Excel.
- Environment note: the writer installed the 'formulas' pip package and says it is not needed. Also, a scratch audit script in the shared scratchpad was overwritten during this run, probably by another agent. No site files were affected; a temporary audit page I placed in the site root was deleted right away.

## T1_scheduling_apps

- 7shifts pricing changed three times between September and November 2026 (Essentials was removed, then restored). The page prints Essentials at $39.99 per location per month for teams up to 30 and Pro at $89.99 for 31+ employees, which is what the live pricing screenshot shows on annual billing. 7shifts' plans page shows Pro at $79.99 and Premium at $134.99 for smaller teams. Re-check 7shifts.com/pricing at each team size before publishing.
- 7shifts free Comp plan cap of 15 employees comes from 7shifts' own plans page. Some third-party sites say 20 or 30.
- Homebase prices ($30, $70, $120 per location) sit under an Annual/Monthly toggle that does not say which billing period they are. Confirm whether they are the annual or the monthly rate.
- When I Work's $2.50, $5 and $8 per-user prices are not labeled monthly or annual, and the Time Tracking & Attendance add-on price is not published.
- ZoomShift Starter is printed as '$2 to $2.50 per active user/month' because the pricing page shows both figures unlabeled. Confirm which is annual and which is monthly.
- Deputy's trial is 'up to 31 days' per its home page (its help center says you add days by completing tasks). Deputy is also running a promotion through November 30, 2026 that the page does not mention.
- Square Shifts: the 30-day trial and the $49 per location Square Plus price come from Square's own FAQ and plan copy. Square's pricing page showed a placeholder instead of the trial length.
- Sling's free plan includes open shifts, shift offers, time-off requests and private messaging. Swaps, unavailability and group messaging are Premium, and shift tasks are Business only. This is from the getsling.com/pricing comparison table; confirm if Sling changes its plans.
- 7shifts compliance features (Chicago, Seattle and Oregon predictive scheduling, minor limits against a school calendar, California breaks) come from its help-center release notes and were not tested.
- HotSchedules is shown with Fourth's logo and a screenshot that still has a regional-site banner and a cookie bar. Replace both with a HotSchedules press-kit image if one becomes available.
- The Workforce.com screenshot (Fair Workweek page) still shows a small '8,134 reviews' star badge under the buttons. Decide whether that is acceptable or recapture another page. The home-page and scheduling-software captures either showed a 4.8 rating with customer logos or timed out.
- Legal notes to re-check each year: Chicago Fair Workweek thresholds and the current pay ceiling for covered employees; NYC fast food 14 days and retail 72 hours; Oregon 14 days and 10-hour rest for employers with 500+ employees; the federal $2.13 tipped cash wage and the rule that tipped overtime is based on the full minimum wage; federal hours limits for 14- and 15-year-olds; California meal and rest premiums. 'Holiday premiums are a policy choice in most states' was not checked state by state.
- Several internal links point to planned pages that are not built yet (for example /blog/how-to-calculate-overtime/, /blog/gps-time-clock-app/, /blog/dupont-schedule/, /blog/holiday-pay/ and the /features/ pages). They are all in plan.json but will return 404 until those pages are published.

## T2_time_clock_apps

- The writer's report that /og-image.png shows a rupee payslip is out of date. The current file shows 'Next paycheck $1,284.50' in USD, so no change is needed.
- When I Work: the pricing page does not show the price with Time Tracking & Attendance switched on; a July 2026 third-party source puts it at $4/$7/$10 per user. Confirm the real time-clock price before publishing, or keep the 'time tracking extra' wording.
- Jibble: its pricing and upgrade pages list only the free plan and do not say which features are free. GPS and facial recognition on the free plan rest on Jibble's help center; Jibble marketing elsewhere mentions 2 free geofences. Confirm before publishing.
- Connecteam: the pricing page's comparison grid is unclear on whether GPS and geofencing are on the free Small Business plan. Its claim that the free plan includes 'all hubs and features' supports the GPS-stamp wording.
- Homebase: the pricing-page checkmarks did not render as text. The plan for location-verified mobile clock-in (Essentials) and the plan for overtime alerts should be confirmed in a browser.
- Clockify free plan: the card says 'up to 5 users' but the FAQ says you can invite 4 people. The pages use 5.
- Deputy: the US pricing page does not state the trial length. 'Up to 31 days, no credit card' comes from the Deputy home page screenshot and its AU pricing page. Geofencing tier is not stated.
- Buddy Punch blocks automated fetches (403). Prices ($4.99/$6.99/$12.99 annual + $19 base), Pro-only geofencing and free admins were confirmed only through search snippets of buddypunch.com/pricing. Check in a browser.
- ClockShark: prices are filled in by JavaScript and were confirmed only from the rendered screenshot. Its pricing page says 'Contract term of three (3) years applies to all pricing plans'; keep that wording exact.
- QuickBooks Time: the current pricing page shows 50%-off promos for 3 months and lists geofencing under Workforce Elite, while Intuit's geofencing page says Time Elite only. The pages use list prices and Time Elite.
- Hubstaff: the Locations add-on price is not published (third parties say $3.33/seat), and monthly prices are not shown.
- Harvest: its free plan card lists only 'Mac & iOS apps', while its feature table shows web, desktop and mobile on Free.
- TimeCamp: trial length is not stated, and the JSON-LD prices differ from the visible table (the visible table is used).
- OnTheClock: its own calculator ($65 for 2 employees) appears to include payroll even with payroll toggled off; the listed $5 + $4/employee is used.
- Legal notes written from general knowledge, to confirm on official sources: electronic-monitoring notice laws in NY (Civil Rights Law §52-c), CT (§31-48d) and DE (19 Del. C. §705); FLSA travel-time rules (29 CFR 785.38); mileage reimbursement (California Labor Code §2802); fair workweek jurisdictions (NYC, SF, Seattle, Chicago, Philadelphia, Oregon); California 30-minute meal period over 5 hours and daily overtime after 8; FLSA record retention (3 years payroll, 2 years time cards).
- Toggl Track and Buddy Punch screenshots have cookie banners across the bottom quarter to third; headlines are visible and the Toggl alt mentions the notice. Re-capture if a cleaner image is wanted.
- Process note: the session scratchpad is shared with other sessions. My early scratch files (audit.py, alt.txt, logos.png, shots0-4.jpg) may have overwritten files of the same name from other sessions, and one of mine was overwritten in turn. Later work went into scratchpad/aud_timeclock5/.

## T5_alternatives

- 7shifts pricing changes often. pricingsaas.com logged a change in late September 2026 and a reversal in early October. Today's site code shows two price sets. For 1 to 30 employees (annual): Essentials $39.99, Pro $79.99, Premium $134.99 (monthly $44.99, $89.99, $149.99). The 31+ employees tier is the page default: Pro $89.99, Premium $179.99 (shown as $99.99 and $199.99 in the comparison table). Check this on 7shifts.com/pricing before publishing.
- When I Work's pricing page contradicts itself. The plan cards show a Time Tracking & Attendance toggle with no price, but the FAQ says every plan includes time and attendance. The cards also list multi-location scheduling under Essentials ($2.50), while the FAQ says multi-location plans cost $5 per user. All four pages say time tracking is an add-on and that Essentials includes multi-location scheduling. Confirm both points with When I Work.
- Homebase annual prices ($24, $56, $96) are worked out from the monthly prices ($30, $70, $120) and the 'Save 20%' badge. Only one set of prices is shown on the page.
- Deputy's 'up to 31 days, no credit card' trial comes from the Deputy home page, not the pricing page. The pricing page says annual billing saves 10% but does not show annual prices.
- Buddy Punch: the $4.99 Starter and $6.99 Pro annual prices (monthly $5.99 and $7.99), the $19 base fee, the $1 scheduling add-on on Starter and the $2 real-time GPS add-on are confirmed on its pricing page today. Its prices have changed several times in the past. The Buddy Punch screenshot shows the vendor's own '4.8 (1,600+ reviews)' badge and a cookie banner. Every Buddy Punch page I captured shows the same badge. Decide whether a screenshot showing a vendor's rating is acceptable under the no-ratings rule.
- Gusto: Simple is $49 plus $6, Plus is $80 plus $12 (with time tracking and scheduling), and Time & Attendance Plus is $6 per person on Simple. These were read from gusto.com/product/pricing in headless Chrome, because WebFetch was blocked with a 403.
- QuickBooks Time: the 30-day free trial covers only the monthly base fee, not per-user fees. The 50% base-fee discount for 3 months is a promotion.
- Connecteam's free plan limits (one time clock, one schedule, no payroll integration) come from its comparison table. Check that this wording is still current.
- Platforms (web, iOS, Android) for 7shifts, Gusto, Homebase web and Deputy are widely known but were not each confirmed on a vendor page.
- Legal notes have not been checked against dol.gov or the city and state sites. They cover predictive scheduling (Oregon, Seattle, San Francisco, NYC, Chicago, Philadelphia), California daily overtime after 8 hours, and FLSA record keeping (3 years for payroll records, 2 years for time cards).
- The When I Work logo now comes from the official SVG and the 7shifts home.jpg was captured manually. Both folders are shared with other listicles, so those pages now use these assets too.
- The other agent's listicles outside this batch repeat the same Connecteam 'every feature' claim and the Gusto/QuickBooks Time wording. They were not edited here and should get the same corrections: best-employee-scheduling-apps, best-hr-software-for-small-business, best-hris-for-small-business, best-leave-management-software, best-retail-scheduling-software, best-scheduling-apps-for-small-business, employee-schedule-makers and free-time-clock-apps.

## P3_letter_templates

- Federal exempt salary level of $684/week ($35,568/yr) and California's 2026 exempt minimum of $70,304: confirm before publishing, and note that 2027 state figures change on January 1.
- California Fair Chance Act coverage (5+ employees, criminal history only after a conditional offer): the writer read search snippets only.
- New York pay notice at hire (LS 195.1, six-year retention) and the 5-working-day written termination notice (LS 195.6): checked against statute text only, because dol.ny.gov returned 403.
- Nevada paid leave line in the seasonal offer sample ('Paid leave as required by Nevada law'), and Nevada's immediate final pay on discharge in the end-of-season letter.
- Final pay timing in the Ohio, Georgia, Florida and New York samples (next regular payday): this is sample data and was not checked state by state.
- WARN Act thresholds: confirmed only through DOL Employment Law Guide and elaws search results, not by reading the full current page.
- 'Many states have continuation (mini-COBRA) laws for smaller employers' and 'many states restrict cash-shortage deductions further' are general statements; no specific state was checked.
- NLRB points (concerted activity with or without a union; Weingarten rights) and the OSHA 1904.35 / Section 11(c) anti-retaliation points: based on agency search results, because the page fetches failed.
- FTC noncompete rule: vacated, with appeals dismissed in 2025. Recheck for any new FTC action before publishing.
- Site-wide: /templates/ and /tools/ hub pages, linked from the breadcrumb and footer, do not exist yet.
- Shared kit: tplkit.py's fixed footer overlaps text on full pages. Each build script works around it with a local CSS override. Fix this in tplkit.py itself before the other batches are built.

## P2_time_pay_templates

- Pay stubs: 2026 Social Security wage base $184,500 and rates of 6.2%, 1.45% and 0.9% over $200k; CA SDI 1.3%; NY PFL 0.432% with a $411.91 cap; NY DBL employee share of $0.60 a week. All change every January, so update the page and the build script settings for 2027.
- Pay stubs: Section 125 premiums are excluded from FICA wages and 401(k) deferrals are not (Pub 15-B). Reported cash tips are FICA wages. Neither was fetched from irs.gov this session.
- Pay stubs: Texas allows the federal $2.13 cash wage with a $5.12 tip credit; confirm on twc.texas.gov.
- Pay stubs: the NY LS 49 form edition date. The page does not print a date, but the writer reported 12/25 and search results showed 11/21.
- Timesheets: overtime for a workweek that straddles two semimonthly periods is paid on the payday for the period in which the workweek ends (29 CFR 778.106).
- PTO: the new wording says several states require payout and a few treat earned vacation like California. Confirm before naming any specific states.
- PTO: the team PTO calendar file defaults to the year 2027 while the preview shows 2026. Both are fine, but confirm that is intended.
- Attendance and PTO: state and city paid sick leave protections are written only as general principles, so readers are told to check their state.
- Every page states that the demo takes 20 minutes, as the reference post does. Confirm this is the real demo length.

## T3_payroll_hr

- Patriot Software was removed from all three payroll pages: its site blocks automated access, and I had no real screenshot, logo or prices from its own pricing page. To bring it back, someone needs to capture a screenshot and logo by hand and check the prices ($17 + $4 Basic, $37 + $5 Full Service, $12 per extra state, 30-day trial) on patriotsoftware.com. assets/tools/patriot-software/home.jpg (Cloudflare page) and logo.png (placeholder 'P') are still on disk, and other batches should not use them.
- QuickBooks Workforce per-employee fees conflict between sources. The captured US pricing screenshot shows +$7 and +$13, which the page uses ($50 + $7, $88 + $13, $134 + $17). WebFetch and Gusto's comparison page show $6.50, $10 and $12, and a July 2026 fee change was announced. Confirm the payroll-only prices on Intuit's US plans page.
- Gusto ($49 + $6, $80 + $12, $180 + $22, plus the Next-Day, Remote I-9 and Time & Attendance Plus add-on prices) and Justworks ($50 + $8 Payroll, $79 PEO Basic, $124 PEO Plus, $8 time tracking, $30 HR consulting) both returned 403 to WebFetch. The base prices match third-party sources, but Justworks PEO Plus is listed at $109 on some sites and $124 on others. Check both pricing pages by hand.
- OnPay's pricing page shows a $49 base + $6 per worker but its footer says 'starts at $55'. The pages note this.
- Time-limited promotions that change often: ADP RUN up to 6 months free (countdown to November 26), Homebase 6 months free payroll + Plus, Roll by ADP 3 months free, Paychex 'Do Not Pay Until April 2027' (first payroll by November 24, 2026, one-year agreement), QuickBooks 50% off for 3 months, and Paycor's conflicting 50%-off / free-months offers. Several of these expire within weeks of November 7, 2026.
- Paychex claims not shown on the page I fetched: that clocking in through the Flex app needs a Paychex Flex Time subscription, and the iOS/Android app details.
- Payroll4Free's Windows requirement and time-clock import are only in commented-out HTML on its site, so I removed both. Confirm the current system requirements if you want to say anything about platforms.
- The Rippling screenshot (assets/tools/rippling/home.jpg) was captured from an India IP and shows a small India flag next to 'EN' in the header. Consider recapturing it from a US location. The Zoho Payroll screenshot was cropped to remove an 'India site' banner, so its left and right edges are slightly trimmed.
- Legal and tax lines worth a quick check: the 1099-NEC threshold of $2,000 for payments in 2026 ($600 before); DOL record retention (3 years for payroll records, 2 for time cards) and the IRS 4-year rule for employment tax records; IRS failure-to-deposit penalties of 2% to 15%; and the Gusto-alternatives tip that switching at the start of a calendar quarter keeps Form 941 simple, which is general practice and is not quoted from the IRS.

## P4_people_templates

- Shared problem not fixed because tplkit is outside my scope: the PDF footer 'Free template from workzenly.com' (position:fixed; bottom:0 with @page margin 14mm in content-plan/tplkit.py) overlaps the last line on full pages. Seen on page 1 of performance-improvement-plan-template.pdf ('Employee comments') and page 3 of restaurant-employee-handbook-template.pdf. Fix it in tplkit._CSS by increasing the @page bottom margin or moving the footer into the margin, then rebuild every listicle's PDFs.
- The nav and footer link to /templates/ and /tools/, but no hub index pages exist yet, and both URLs are also in the BreadcrumbList.
- Texas restaurant sample: serving age 18 for on-premise service and TABC seller-server training are written as company policy. Third-party sources disagree on whether certification is legally required, so check tabc.texas.gov.
- Texas food handler card within 60 days. One DSHS FAQ says 30 days, and local health departments may be stricter.
- EEOC record retention (29 CFR 1602.14): one year from the date the record was made or the personnel action, whichever is later, and one year after an involuntary termination. The page's wording is a simplified version of this.
- Colorado daily overtime (over 12 hours a day) and 10-minute rest breaks per 4 hours were confirmed only in 2020 to 2024 COMPS Order summaries. Check the current COMPS Order on cdle.colorado.gov.
- Montana good cause after the probationary period. The page does not state a default probation length.
- California: personnel records kept 3 years after termination and a 30-day inspection response (Labor Code 1198.5), Labor Code 2802 expense reimbursement, and the DE 4 form name.
- I-9: the remote document review alternative procedure applies only to E-Verify employers in good standing (DHS, 2023). Recheck that it is still in effect. Also check that the 01/20/25 Form I-9 edition and the 2026 W-4 are still current.
- Oregon predictive scheduling: covers retail, hospitality and food service employers with 500 or more employees worldwide and requires 14 days' notice.
- New York: pay notice at hire kept 6 years, and a sexual harassment prevention policy given at hire and at each annual training.
- Federal tip credit: $2.13 cash wage and a $5.12 maximum tip credit. Check the current figures on dol.gov.
- Thresholds of 15 employees (Title VII), 20 (ADEA) and 50 (FMLA) were stated from well-established law and not rechecked this session.

## T4_hris_ats_pto

- Gusto prices ($49+$6 Simple, $80+$12 Plus, $180+$22 Premium) and its iOS and Android apps: gusto.com/pricing returned 403 to WebFetch, so these are confirmed only by third-party sources dated October 2026.
- Justworks prices ($50 + $8 per employee Payroll, $79 PEO Basic, $124 PEO Plus) and which phone systems its app runs on: justworks.com returned 403, so these are confirmed only by third-party sources.
- BambooHR 7-day free trial: stated on BambooHR marketing pages found through search, but the pricing page does not give a trial length.
- Homebase prices ($30/$70/$120 per location): the pricing page shows one price per plan with a 'Save 20% annually' toggle, so it is unclear whether these are monthly or annual-billing prices. The pricing page also says Plus is '$0 with payroll promo' and Essentials and Plus include '6 months free with payroll', both time-limited.
- Jibble: its PTO page says 'free forever for unlimited users', but a Jibble help article says time-off balances need Premium or Ultimate. Confirm with Jibble which leave features are free before publishing the free-PTO page.
- Workstream's 'up to 8 months free' offer is a time-limited promotion.
- Recruitee publishes no US-dollar prices or trial length; it is listed as 'Quote'.
- Day Off trial length is not stated on its pricing page.
- Rippling's policies by location and employee type and its templates for state and local rules come from Rippling's own blog and product pages. The base fee is not published. The Rippling screenshot (assets/tools/rippling/home.jpg, its HRIS product page) still shows an India flag in the header because of where it was captured from, so consider recapturing it from a US connection.
- assets/tools/bamboohr/home.jpg is shared with tools/best-hr-software-for-small-business. I replaced a pricing-form capture with a capture of the real home page. Make sure no parallel batch overwrites it again. It still shows a privacy box in the lower left.
- assets/tools/vacation-tracker/logo.png is the vendor's only published icon, a 32px favicon scaled up, so it is slightly soft. assets/tools/justworks/logo.png is also an upscaled favicon wordmark.
- Legal notes to re-check before publishing: these states' pay-transparency job-posting rules (Colorado, California, Washington, New York); the FCRA rule that you need the candidate's written permission before a background check; EEOC 1-year retention of applications and personnel records; FLSA 3-year payroll records; I-9 Section 2 within 3 business days; California sick leave of 1 hour per 30 hours worked and Washington's 1 hour per 40 hours; the state paid family and medical leave programs named (CA, NY, NJ, WA, MA); and the FMLA eligibility notice within 5 business days.