#!/usr/bin/env python3
"""Build the downloads for /templates/job-offer-letter-templates/.

Run from the site root:  python content-plan/template_builds/job-offer-letter-templates.py
Files go to assets/templates/job-offer-letter-templates/. Letters use [brackets] for the parts to fill in;
the filled-in examples live in the page previews.
"""
import sys
import pathlib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import tplkit
# tplkit's fixed .foot overlaps the last text line when a page is full; print the footer in the
# bottom page margin instead (Chrome page-margin boxes). Local to this build; tplkit is unchanged.
tplkit._CSS += (".foot { display: none; }\n@page { margin-bottom: 18mm; @bottom-left { content: "
                "\"Free template from workzenly.com\"; font-family: Arial, sans-serif; font-size: 8pt; color: #888; } }\n")
from tplkit import out, docx_doc, pdf_doc

d = out("job-offer-letter-templates")

ACCEPT = [
    ("h", "Acceptance"),
    ("p", "I accept this offer on the terms above. I understand that my employment is at will and that this "
          "letter is not a contract for any fixed period of employment."),
    ("sign", ["Candidate signature", "For the company"]),
]


def letter(name, title, text, extra=None):
    blocks = [("letter", text)] + (extra if extra is not None else ACCEPT)
    docx_doc(d / f"{name}.docx", title, blocks)
    pdf_doc(d / f"{name}.pdf", title, blocks)


AT_WILL = ("Employment with [Company name] is at will. This means that you or the company may end the employment "
           "relationship at any time, with or without cause or notice. Nothing in this letter is a contract for a "
           "fixed term of employment.")

# 1. Standard job offer letter ------------------------------------------------------------------
letter("standard-job-offer-letter", "Job offer letter", f"""[Company name]
[Street address, City, State ZIP]

[Date]

[Candidate full name]
[Street address, City, State ZIP]

Dear [Candidate first name],

We are pleased to offer you the position of [Job title] at [Company name]. You will report to [Manager name and title] and work at [work location].

Start date: [Start date]
Employment type: [Full-time / Part-time], [non-exempt (paid hourly, eligible for overtime) / exempt (salaried)]
Pay: [$ amount] per [hour / year], paid [weekly / every other week / twice a month] by direct deposit or paper check
Schedule: [Days and hours, or "as scheduled, about __ hours a week"]

Benefits: You will be eligible for [health insurance, dental, vision, 401(k), paid time off] starting [date or waiting period], under the terms of each plan. Details are in the enclosed benefits summary.

This offer depends on you providing documents that prove your identity and authorization to work in the United States for Form I-9 within three business days of your start date. [Add any other conditions, such as a background check.]

{AT_WILL}

Please sign and return this letter by [deadline date] to accept. If you have questions, contact me at [phone] or [email].

We look forward to having you on the team.

Sincerely,

[Name]
[Title]""")

# 2. Hourly (non-exempt) offer letter ------------------------------------------------------------
letter("hourly-employee-offer-letter", "Offer letter: hourly employee", f"""[Company name]
[Address]

[Date]

Dear [Candidate first name],

Thank you for interviewing with us. We are happy to offer you the job of [Job title] at [location name], reporting to [Manager name].

Your details:
- Start date and first shift: [Day, date], [start time] to [end time]
- Pay rate: $[00.00] per hour
- Overtime: This is a non-exempt position. You will be paid 1.5 times your regular rate for hours worked over 40 in a workweek[, and any daily overtime required by state law]. Overtime must be approved by your manager in advance.
- Workweek: [Sunday 12:00 a.m. through Saturday 11:59 p.m.]
- Expected hours: about [__] hours a week. Hours vary with the posted schedule.
- Paydays: [every other Friday], for the pay period ending [the Sunday before]
- Breaks and meals: [Describe meal and rest break policy, following state law]
- Tips: [If applicable: how tips are pooled or paid]
- Uniform or equipment: [What the company provides and what you need to bring]

You will record your hours with [time clock app, kiosk or paper time card] every shift, including meal breaks.

On your first day, please bring documents for Form I-9 (see the list of acceptable documents on uscis.gov) and your bank details if you want direct deposit.

{AT_WILL}

Please reply or sign below by [deadline date] to accept.

Welcome aboard,

[Name]
[Title]""")

# 3. Salaried (exempt) offer letter --------------------------------------------------------------
letter("salaried-exempt-offer-letter", "Offer letter: salaried exempt employee", f"""[Company name]
[Address]

[Date]

Dear [Candidate first name],

On behalf of [Company name], I am pleased to offer you the full-time position of [Job title], reporting to [Manager name and title]. We were impressed by [one specific thing from the interviews] and believe you will make a real difference to [team or goal].

Compensation
- Base salary: $[00,000] per year, paid [twice a month / every other week] ($[0,000.00] per pay period)
- Classification: Exempt. This position is not eligible for overtime pay under the Fair Labor Standards Act[ and state law].
- Bonus: [Describe any bonus plan, target and how it is decided, or delete]
- Salary review: [Annually in month]

Benefits
- Health, dental and vision insurance starting [date], with the company paying [__%] of the employee premium
- [__] days of paid time off per year, accruing [per pay period], plus [__] paid holidays
- 401(k) plan [with a match of __%] after [eligibility period]
- [Other benefits: phone stipend, training budget, commuter benefits]

Start date and location: [Start date], at [office address / hybrid: __ days a week in the office]

Conditions: This offer depends on [proof of identity and work authorization for Form I-9 / a background check / signing our confidentiality agreement].

{AT_WILL}

This offer is open until [deadline date]. To accept, sign below and return a copy to me.

Sincerely,

[Name]
[Title]""")

# 4. Part-time offer letter ----------------------------------------------------------------------
letter("part-time-employee-offer-letter", "Offer letter: part-time employee", f"""[Company name]
[Address]

[Date]

Dear [Candidate first name],

We are glad to offer you a part-time position as [Job title] at [location], reporting to [Manager name].

- Start date: [Start date]
- Pay: $[00.00] per hour, paid [every other week]
- Hours: [__ to __] hours a week, usually [days and times you discussed, such as weekday evenings and every other Saturday]. Hours are not guaranteed and follow the posted schedule.
- Availability: You told us you are available [availability]. Please tell your manager in writing at least [two weeks] before your availability changes.
- Overtime: If you work more than 40 hours in a workweek, you will be paid 1.5 times your regular rate.
- Benefits: Part-time employees are eligible for [paid sick time as required by law / employee discount / 401(k) after __ hours]. [Not eligible for: list.]

Please bring documents for Form I-9 on your first day.

{AT_WILL}

Please sign and return this letter by [deadline date].

Best regards,

[Name]
[Title]""")

# 5. Conditional (contingent) offer letter ------------------------------------------------------
letter("conditional-job-offer-letter", "Conditional job offer letter", f"""[Company name]
[Address]

[Date]

Dear [Candidate first name],

We are pleased to make you a conditional offer of employment for the position of [Job title] at [location], starting [planned start date], at $[pay] per [hour / year].

This offer is conditional on the following. We will not finalize your start date until each one is complete:
1. A background check by [screening company], which we will run only after you sign the separate disclosure and authorization form we are sending with this letter
2. [A drug test at an approved lab within __ days, if allowed by law for this role]
3. [Verification of your license or certification: type and number]
4. [Satisfactory references from __ previous employers]
5. Documents that prove your identity and authorization to work in the United States for Form I-9, no later than three business days after your first day of work

If something in a background report might cause us to withdraw this offer, we will first give you a copy of the report and a summary of your rights, and give you the chance to respond before we make a final decision.

All other terms:
- Reports to: [Manager name and title]
- Classification: [Non-exempt hourly / Exempt salaried]
- Schedule: [schedule]
- Benefits: [summary or "see enclosed summary"]

{AT_WILL}

Please sign below by [deadline date] to accept this conditional offer and return the authorization form.

Sincerely,

[Name]
[Title]""")

# 6. Internship offer letter --------------------------------------------------------------------
letter("internship-offer-letter", "Internship offer letter", f"""[Company name]
[Address]

[Date]

Dear [Candidate first name],

We are excited to offer you a paid internship as [Internship title] on our [team] team for [season and year].

- Dates: [Start date] to [End date] ([__] weeks)
- Supervisor: [Name and title]
- Pay: $[00.00] per hour, paid [every other week]. Overtime at 1.5 times your regular rate applies to hours over 40 in a workweek.
- Hours: [__] hours a week, [days and times]. [Remote / on-site at address / hybrid.]
- What you will work on: [two or three real projects]
- What you will learn: [skills, training, mentoring]
- [School credit: If your school requires paperwork for credit, send it to your supervisor and we will complete it.]

The internship ends on [End date]. It does not include a promise of a permanent job, although we will consider strong interns for open positions.

Please bring documents for Form I-9 on your first day.

{AT_WILL}

Please sign and return this letter by [deadline date].

Best,

[Name]
[Title]""")

# 7. Seasonal or temporary offer letter --------------------------------------------------------
letter("seasonal-temporary-offer-letter", "Offer letter: seasonal or temporary employee", f"""[Company name]
[Address]

[Date]

Dear [Candidate first name],

We are happy to offer you a seasonal position as [Job title] at [location] for the [season] season.

- Expected dates: [Start date] to about [Expected end date]. The end date may move earlier or later depending on business needs; we will tell you as soon as we know.
- Pay: $[00.00] per hour[, plus a $__ per hour shift differential for __ shifts]
- Overtime: 1.5 times your regular rate for hours over 40 in a workweek. During [peak dates], expect up to [__] hours a week.
- Shifts: [shift times], scheduled [__] days in advance through [scheduling app or posted schedule]
- Training: [paid orientation date and time]
- Safety gear: [what is provided, and what to wear, such as closed-toe safety shoes]
- Benefits: [Paid sick time as required by state or local law / employee discount / not eligible for health plan]

Seasonal employment ends on the last day of the assignment unless we offer you, in writing, a regular position.

Please bring documents for Form I-9 to orientation.

{AT_WILL}

Please sign and return by [deadline date].

Thank you,

[Name]
[Title]""")

# 8. Remote employee offer letter --------------------------------------------------------------
letter("remote-employee-offer-letter", "Offer letter: remote employee", f"""[Company name]
[Address]

[Date]

Dear [Candidate first name],

We are pleased to offer you the fully remote position of [Job title], reporting to [Manager name and title].

- Start date: [Start date]
- Work location: You will work from your home in [City, State]. Please tell us before you move to another state, because your work state decides which payroll taxes and employment laws apply.
- Pay: $[amount] per [hour / year], paid [frequency]. Classification: [non-exempt, overtime eligible / exempt].
- Working hours: [core hours and time zone], [__] hours a week. [Non-exempt: record all time worked, including time spent on email or chat outside your shift.]
- Equipment: We will ship you [laptop, headset, monitor] before your start date. It remains company property and must be returned when your employment ends.
- Expenses: We will [reimburse internet and phone costs up to $__ per month / pay a $__ monthly stipend].
- Travel: [Expected in-person meetings per year, paid by the company.]
- Benefits: [summary]

Before your first day we will send you remote onboarding forms. For Form I-9 we will [arrange an in-person document review / use the remote document examination process available to E-Verify employers] within three business days of your start date.

{AT_WILL}

Please sign below by [deadline date].

Sincerely,

[Name]
[Title]""")

# 9. Job offer email and acceptance reply ------------------------------------------------------
OFFER_EMAIL = """Subject: Job offer: [Job title] at [Company name]

Hi [Candidate first name],

Thank you for meeting with us. I am happy to offer you the position of [Job title] at [Company name], starting [start date].

The highlights:
- Pay: $[amount] per [hour / year], paid [frequency]
- Schedule: [schedule]
- Reports to: [manager]
- Benefits: [one-line summary]

Your formal offer letter is attached. Please review it, and sign and return it by [deadline date]. I am happy to answer any questions by phone at [phone].

We hope you will join us.

[Name]
[Title]"""

ACCEPT_EMAIL = """Subject: Re: Job offer: [Job title] at [Company name]

Hi [Hiring manager first name],

Thank you for the offer. I am pleased to accept the position of [Job title] at $[amount] per [hour / year], starting [start date]. My signed offer letter is attached.

Please let me know what I should bring or complete before my first day.

Best regards,
[Your name]
[Phone]"""

COUNTER_EMAIL = """Subject: Re: Job offer: [Job title] at [Company name]

Hi [Hiring manager first name],

Thank you for the offer. I am excited about the role and the team. Before I accept, I would like to ask whether you can move the starting pay to $[amount] per [hour / year], based on [specific reason: certification, years of experience, current pay].

If that is possible, I am ready to sign and can start on [date]. Thank you for considering it.

Best regards,
[Your name]"""

email_blocks = [("h", "1. Job offer email (employer to candidate)"), ("letter", OFFER_EMAIL),
                ("h", "2. Acceptance reply (candidate to employer)"), ("letter", ACCEPT_EMAIL),
                ("h", "3. Counteroffer reply (candidate to employer)"), ("letter", COUNTER_EMAIL)]
docx_doc(d / "job-offer-email-templates.docx", "Job offer email templates", email_blocks)
pdf_doc(d / "job-offer-email-templates.pdf", "Job offer email templates", email_blocks)

for f in sorted(d.iterdir()):
    print(f"{f.name:50s} {f.stat().st_size:>8,d} bytes")
