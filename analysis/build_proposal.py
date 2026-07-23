"""Builds deliverables/Clipboard_Case_Proposal.pdf (6-page narrative + appendix)."""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate, Frame, NextPageTemplate, PageBreak, PageTemplate,
    Paragraph, Spacer, Table, TableStyle,
)

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables" / "Clipboard_Case_Proposal.pdf"

INK = colors.HexColor("#1a1a2b")
ACCENT = colors.HexColor("#8a3033")
GRID = colors.HexColor("#c9c9d1")
HEADBG = colors.HexColor("#f0eef2")

body = ParagraphStyle("body", fontName="Times-Roman", fontSize=10.2, leading=13.6,
                      textColor=INK, alignment=TA_JUSTIFY, spaceAfter=7)
h1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=13, leading=16,
                    textColor=ACCENT, spaceBefore=13, spaceAfter=6)
h2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=10.6, leading=13,
                    textColor=INK, spaceBefore=9, spaceAfter=4)
title_st = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=17, leading=21, textColor=INK, spaceAfter=2)
sub_st = ParagraphStyle("s", fontName="Helvetica", fontSize=10.5, leading=14, textColor=colors.HexColor("#55555f"), spaceAfter=10)
bullet = ParagraphStyle("bl", parent=body, leftIndent=16, bulletIndent=5, spaceAfter=4)
tcell = ParagraphStyle("tc", fontName="Helvetica", fontSize=8.6, leading=11, textColor=INK)
tcell_b = ParagraphStyle("tcb", parent=tcell, fontName="Helvetica-Bold")
cap = ParagraphStyle("cap", fontName="Helvetica-Oblique", fontSize=8.4, leading=10.5,
                     textColor=colors.HexColor("#55555f"), spaceBefore=3, spaceAfter=8)


def tbl(rows, widths, header=True):
    data = [[Paragraph(c, tcell_b if (header and i == 0) else tcell) for c in r] for i, r in enumerate(rows)]
    t = Table(data, colWidths=widths, hAlign="LEFT")
    style = [
        ("GRID", (0, 0), (-1, -1), 0.5, GRID),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), HEADBG))
    t.setStyle(TableStyle(style))
    return t


def p(txt, style=body):
    return Paragraph(txt, style)


def build():
    S = []
    S.append(p("Reducing the cost of late cancellations by winning the refill race", title_st))
    S.append(p("Project proposal for executive leadership &nbsp;|&nbsp; David Ezieshi &nbsp;|&nbsp; "
               "Operations Strategy &nbsp;|&nbsp; February 1, 2022 &nbsp;|&nbsp; "
               "Data: Cleveland, shift starts Oct 1, 2021 – Jan 31, 2022", sub_st))

    # ------------------------------------------------------------------ 1
    S.append(p("1. Summary of the recommendation", h1))
    S.append(p(
        "Late cancellations are usually framed as a discipline problem: workers bail on commitments, so we "
        "should tighten the screws. We tried that. The 2021 attendance policy moved the needle only "
        "incrementally, and the data explains why: the top reported reasons for late cancellation — sickness, "
        "family emergencies, car trouble — are mostly legitimate and unpreventable. We cannot deactivate our "
        "way past a sick child, and every unit of added punishment presses on the flexibility that brings "
        "professionals to Clipboard in the first place. Both sides are our customers."))
    S.append(p(
        "The data supports a different framing. <b>Facilities are not damaged by cancellations; they are "
        "damaged by the shifts that stay empty afterwards.</b> When a professional cancels with 24 or more "
        "hours of notice, the marketplace already recovers: 67% of those shifts get re-claimed and worked. "
        "When the cancellation comes inside 24 hours, recovery collapses to 30%, and after a no-call-no-show "
        "(NCNS) to 14%. Of the 2,634 Cleveland shifts in this window that a professional abandoned and no one "
        "worked, 89% trace to late cancellations and NCNS. That failure mode destroyed roughly $730K in "
        "facility bookings (about $160K of our revenue at a 22% take) in four months, in one market, and "
        "touched 55 of our 66 active Cleveland facilities — precisely the experience we know correlates with "
        "facility churn."))
    S.append(p(
        "<b>I recommend we invest in one thing: winning the refill race.</b> When a booked shift fails — a "
        "late cancellation or a no-show — Clipboard should detect the failure within minutes and re-offer the "
        "shift instantly to professionals who have proven they claim same-day work. The marketplace is "
        "already telling us this works: shifts rescued by a same-day claim after a cancellation get worked "
        "80% of the time, indistinguishable from shifts that were never cancelled at all. Today those rescues "
        "happen by luck. I propose we make them happen by design."))
    S.append(p(
        "The 90-day target is to cut cancel-driven empty shifts in Cleveland by ~104 per month — closing "
        "half the recovery gap between late and early cancellations — worth roughly $85K a year in retained "
        "revenue in Cleveland alone before any churn effect, with a mechanism that generalizes to every "
        "market and asks nothing punitive of our professionals."))

    # ------------------------------------------------------------------ 2
    S.append(p("2. Diagnosis: the harm is concentrated in the recovery gap", h1))
    S.append(p(
        "I treated the shifts log as the universe (35,926 shifts after removing 5,114 junk rows with "
        "non-positive charge or duration) and joined booking and cancellation events onto it. Definitions, "
        "cleaning steps, and every table below are reproduced in the attached Excel models; assumptions are "
        "registered in Appendix D."))
    S.append(p("2.1 Recovery depends almost entirely on notice", h2))
    S.append(tbl([
        ["Final cancel event on shift", "Shifts", "Refilled &amp; worked", "Died empty", "Deleted by facility"],
        ["Early cancel (≥24h notice)", "2,392", "67.4%", "11.8%", "20.8%"],
        ["Late cancel (&lt;24h notice)", "2,288", "30.2%", "60.1%", "9.7%"],
        ["No-call-no-show", "1,170", "13.8%", "83.4%", "2.7%"],
    ], [1.85 * inch, 0.7 * inch, 1.25 * inch, 0.9 * inch, 1.35 * inch]))
    S.append(p("Each shift counted once, classified by its final cancellation event. Source: models workbook, Calc_ShiftOutcomes tab.", cap))
    S.append(p(
        "Early cancellations are a solved problem — two-thirds get re-claimed and worked with no intervention "
        "from us. The damage is concentrated where notice runs out: late cancellations and NCNS produce "
        "1,376 and 976 empty shifts respectively — 89% of all cancel-driven empties. (A further 6,936 clean "
        "shifts went unworked without ever being claimed or cancelled; that is a demand-fill problem outside "
        "this case's scope, and none of the figures here depend on it.)"))
    S.append(p("2.2 Within the danger zone, speed is what separates recovery from loss", h2))
    S.append(tbl([
        ["Time remaining when late cancel lands", "Shifts", "Refilled &amp; worked", "Died empty"],
        ["12–24 hours", "316", "44.9%", "37.3%"],
        ["4–12 hours", "669", "33.5%", "57.2%"],
        ["Under 4 hours", "1,303", "24.9%", "67.2%"],
    ], [2.3 * inch, 0.7 * inch, 1.3 * inch, 1.0 * inch]))
    S.append(p("Source: models workbook, Calc_ShiftOutcomes tab.", cap))
    S.append(p(
        "Recovery odds nearly double when there is half a day of runway. The median late cancellation "
        "leaves just 3.1 hours before shift start, so every minute between the cancellation event and the "
        "re-offer is a minute we cannot afford. NCNS is worse in a specific, fixable way: the median NCNS is "
        "logged roughly <b>35 hours after the shift started</b>. We are not losing the refill race on "
        "no-shows — we are not finding out there was a race until the day after."))

    # ------------------------------------------------------------------ 3
    S.append(p("3. Three tempting answers the data rules out", h1))
    S.append(p("3.1 Punish late cancellations harder", h2))
    S.append(p(
        "Beyond the policy already being live and the reasons being mostly legitimate, escalation has a "
        "perverse failure mode: a professional who cannot cancel without penalty can still simply not show "
        "up. Every late cancellation converted into a no-show moves a shift from a 60% chance of dying empty "
        "to an 83% chance, and moves the facility from some warning to none. Harsher cancellation penalties "
        "risk manufacturing our worst outcome."))
    S.append(p("3.2 Predict who will flake and intervene beforehand", h2))
    S.append(p(
        "I tested this directly with a point-in-time analysis: for each of 9,713 bookings, I counted the "
        "professional's prior late cancellations and NCNS strictly before the moment of booking, then "
        "checked whether that same professional went on to fail that booking. History does carry signal at "
        "the extreme — professionals with 3+ prior offenses fail 23.6% of subsequent bookings versus a 13.0% "
        "base (1.8x) — and I will use that flag. But it cannot be the strategy, because <b>42% of no-shows "
        "are committed by professionals with zero prior offenses of any kind, and 70% by professionals with "
        "no prior no-show</b>. Most of the worst events come from people no history-based system can see "
        "coming. Prevention is structurally capped; recovery is not."))
    S.append(p("3.3 Add friction to last-minute bookings (and why I killed my own best finding)", h2))
    S.append(p(
        "My most striking early finding: shifts claimed within 24 hours of start end in failure 31.3% of the "
        "time, twice the 15.4% rate of shifts booked ahead. The obvious move is to add commitment friction "
        "to last-minute claims. Before proposing it, I re-ran the test at the level that actually matters — "
        "does <i>that professional</i> cancel <i>that booking after making it</i> — and the finding "
        "inverted. Last-minute claimers abandon only 12% of their claims, versus 26% for professionals who "
        "book days ahead (life intervenes in the interim). The 31.3% was reverse causality: 27% of "
        "last-minute claims are made on shifts <i>someone else already cancelled</i> — the claim is the "
        "rescue, not the risk. Shifts saved by such rescues get worked 80.5% of the time. Same-day claimers "
        "are not our most fragile users; they are our recovery engine, and adding friction to them would "
        "sabotage the exact behavior this proposal scales. Full test in Appendix C."))

    # ------------------------------------------------------------------ 4
    S.append(p("4. The solution: a rescue loop that makes recovery the default", h1))
    S.append(p(
        "One solution, stated in one sentence: <b>when a booked shift fails, Clipboard detects it within "
        "minutes and re-offers it instantly to proven same-day claimers, so the shift is worked instead of "
        "dying empty.</b> Three jobs-to-be-done make up this single pipeline — none is a separate "
        "initiative, and none stands alone:"))
    S.append(p(
        "<b>JTBD 1 — Detect the failure the moment it happens.</b> Late cancellations already hit our logs at "
        "t=0; today nothing operational fires. No-shows we learn about a median of 35 hours late. I will add "
        "a one-tap “professional hasn't arrived” report on the facility side, prompted automatically 15 "
        "minutes after shift start when the professional has not confirmed arrival in the app. Detection "
        "target: late cancels trigger the loop in under 5 minutes; half of NCNS reported within 1 hour of "
        "start (vs. ~2% today within that hour).", bullet))
    S.append(p(
        "<b>JTBD 2 — Re-offer instantly to the same-day pool.</b> A failed shift goes to the top of a "
        "dedicated Urgent Shifts feed and triggers targeted push notifications to professionals filtered on "
        "three proven criteria: matching license, history of same-day claims, and history with that facility. "
        "This pool is real, not hypothetical: 4,259 distinct professionals made a sub-24-hour claim in the "
        "booking log, 356 of them in Cleveland in this window, and their claims hold 88% of the time. At "
        "target, we need them to absorb roughly 3–4 additional rescues a day across the whole market.", bullet))
    S.append(p(
        "<b>JTBD 3 — Start the race early for the shifts we can see coming.</b> The one predictive signal "
        "that survived testing — 3+ prior offenses, 1.8x failure risk, 16% of bookings — aims the same "
        "pipeline earlier. Flagged bookings get a confirmation request 24 hours before start. No response "
        "does not punish the professional and does not release the shift; it quietly surfaces the shift to "
        "the Urgent feed as “backup wanted,” so if the cancellation comes, the race starts with hours of "
        "runway instead of minutes. This is a targeting input to the rescue loop, not a second solution.", bullet))
    S.append(p(
        "What this deliberately does <i>not</i> do: it adds no penalty, no booking friction, and no new "
        "obligation for professionals — cancelling stays exactly as flexible as today, and the existing "
        "attendance policy stays exactly as it is. The intervention consumes an event stream we already have "
        "and routes demand to supply that already behaves this way. That is why it is cheap: roughly two "
        "engineer-months (event trigger, feed surface, push targeting, facility one-tap) plus ops time I "
        "will staff myself."))
    S.append(p(
        "The honest limit, stated plainly: the NCNS residual. No-shows are 37% of cancel-driven empties, "
        "they are mostly first offenses no signal can flag, and detection-then-rescue can usually save only "
        "part of the shift's hours. This proposal recovers late cancellations well, recovers no-show hours "
        "partially, and does not pretend otherwise. I accept that residual rather than chase it with "
        "punitive tools the data says will backfire."))

    # ------------------------------------------------------------------ 5
    S.append(p("5. Metrics: how we will know, by when", h1))
    S.append(p(
        "Baseline honesty first: late-cancel refill improved through the window (14% in Oct 2021 to 39% in "
        "Jan 2022) as the marketplace matured, so I benchmark against the most recent month, not the "
        "flattering 4-month average, and I track the <i>gap</i> to early-cancel refill so marketplace-wide "
        "drift nets out of our scorecard."))
    S.append(tbl([
        ["Metric", "Baseline (Jan 2022)", "Day-90 target", "Success / failure line"],
        ["North star: cancel-driven empty shifts per week (Cleveland)", "~134", "≤110",
         "Success ≤115; failure &gt;125 → kill or redesign"],
        ["Primary driver: late-cancel refill-to-worked, 30-day rolling", "39.3%", "≥53%",
         "Success ≥50%; failure &lt;45%"],
        ["Gap to early-cancel refill benchmark (67.4%)", "28.1 pts", "≤14 pts", "Controls for organic drift"],
        ["Median time from failure event to re-offer", "no pipeline (n/a)", "&lt;5 min", "Leading indicator, weekly"],
        ["NCNS reported within 1h of shift start", "~2%", "≥50%", "Leading indicator, weekly"],
    ], [2.15 * inch, 1.25 * inch, 1.0 * inch, 1.9 * inch]))
    S.append(p("Guardrails, reviewed weekly with kill authority attached:", h2))
    S.append(p(
        "<b>Moral hazard:</b> if instant recovery makes cancelling feel consequence-free, late-cancel rate "
        "per 100 booked shifts will rise. Alert at a 10% relative increase; sustained breach pauses rollout. "
        "<b>Supply cannibalization:</b> rescues must be net-new fills, not claims shuffled from other open "
        "shifts — which is exactly why the north star is total empty shifts, not the refill rate alone. "
        "<b>Professional experience:</b> push volume capped per professional per day; confirmation requests "
        "only on flagged bookings (16%); flag never visible to facilities and never attached to a penalty. "
        "<b>Facility trust:</b> one-tap no-show reports are auditable against timesheet data to prevent "
        "misuse."))

    # ------------------------------------------------------------------ 6
    S.append(p("6. Execution: the first 90 days", h1))
    S.append(tbl([
        ["When", "What ships", "Owner"],
        ["Weeks 1–2 (Feb 1–14)", "Metrics instrumented from existing logs; baseline dashboard live; "
         "auto re-list trigger on late-cancel events (no more manual gap)", "Me + 1 backend eng"],
        ["Weeks 3–4", "Urgent Shifts feed + targeted push to same-day claimers, Cleveland pilot; "
         "start with the 10 facilities carrying 50% of late/NCNS empties", "Me + mobile eng (part-time)"],
        ["Weeks 5–6", "Facility one-tap no-show report + arrival confirmation prompt", "Me + facility CS lead"],
        ["Weeks 7–8", "T-24h confirmation requests on 3+ offense bookings; pre-warm flow live", "Me"],
        ["Weeks 9–13", "Weekly metric/guardrail reviews; iterate targeting; Day-90 readout with "
         "scale-to-all-markets or kill decision, pre-committed to the thresholds above", "Me, reporting to exec sponsor"],
    ], [1.15 * inch, 3.85 * inch, 1.3 * inch]))
    S.append(p(
        "If refill conversion stalls below target with the pipeline working (offers going out in minutes but "
        "not converting), the diagnosis is incentive, not speed, and the next lever is a same-day claim "
        "bonus funded by the ~$75 average take per rescued shift. I am deliberately not bundling that spend "
        "into this proposal; it is the pre-identified second iteration, triggered by the data."))
    S.append(p(
        "The prize at target: ~104 additional shifts worked per month in Cleveland (~$85K/year of protected "
        "revenue in one market), the same mechanism ready for every other market — and, more valuable than "
        "either, facilities who see failed shifts recovered instead of experiencing empty ones. That is the "
        "outcome that keeps them on the platform."))

    # ------------------------------------------------------------------ appendix
    S.append(NextPageTemplate("appendix"))
    S.append(PageBreak())
    S.append(p("Appendix", title_st))
    S.append(p("Supporting analysis. The attached Excel models (Clipboard_Analysis_Models.xlsx) carry the shifts "
               "log and per-row analysis tabs (each classification and join shown as a column), feeding "
               "live-formula summary and sorted tabs — every figure can be traced, filtered, and re-sorted in "
               "the workbook; tab references given per table.", sub_st))

    S.append(p("A. Data cleaning and definitions", h1))
    S.append(p(
        "<b>Universe:</b> 41,040 raw shifts; 5,114 rows removed for non-positive charge or duration (data "
        "errors); 35,926 clean shifts across 66 facilities, Oct 1 2021 – Jan 31 2022 starts; 91.5% CNA/LVN. "
        "The shifts table is the anchor; booking (127,005 claims) and cancellation (78,073 events) logs are "
        "joined onto it — both logs cover a wider date range than the shifts window, which is why raw join "
        "rates are below 100%. <b>Worked</b> = Verified timesheet. <b>Facility-deleted</b> = Deleted flag "
        "(cancelled by facility; excluded from harm counts). <b>Empty</b> = neither worked nor "
        "facility-deleted. <b>Early / late cancel</b> = WORKER_CANCEL with ≥24h / &lt;24h lead. <b>NCNS</b> "
        "= NO_CALL_NO_SHOW action. Shifts with multiple cancel events (a shift can be claimed, cancelled, "
        "and re-claimed) are classified once, by the final event, so every table sums without double "
        "counting: 6,960 cancel events on the clean universe collapse to 5,850 distinct cancelled shifts."))

    S.append(p("B. Key volumes", h1))
    S.append(tbl([
        ["Quantity", "Value", "Workbook tab"],
        ["Cancel events on clean universe (early / late / NCNS)", "3,347 / 2,436 / 1,177", "Calc_Notice"],
        ["Notice given, worker cancels: ≥72h / 24–72h / 4–24h / &lt;4h", "48.8% / 9.1% / 18.7% / 23.4%", "Calc_Notice"],
        ["Cancel-driven empty shifts (late / NCNS / early)", "1,376 / 976 / 282 = 2,634", "Calc_ShiftOutcomes"],
        ["Facility charge destroyed by late+NCNS empties (4 mo)", "$729,559 gross; ~$160K CBH take", "Calc_Prize"],
        ["Facilities hit by ≥1 late/NCNS empty; top-10 share", "55 of 66; 50.2%", "Sort_Facilities"],
        ["NCNS concentration: top 10% of NCNS workers", "30.9% of NCNS events", "Sort_Workers_NCNS"],
    ], [3.0 * inch, 2.2 * inch, 1.2 * inch]))

    S.append(p("C. The prediction test (point-in-time, leakage-free)", h1))
    S.append(p(
        "Method: for each booking on the clean universe (n=9,713), offense history is counted strictly "
        "before the booking timestamp; the outcome is whether the same professional late-cancelled or "
        "NCNS'ed that same booking after making it. This removes two contaminations present in a naive "
        "shift-level analysis: failures that predate the claim (rescue claims), and other professionals' "
        "failures on the same shift."))
    S.append(tbl([
        ["Prior late/NCNS offenses at claim", "Bookings", "Fail % (after booking)", "Lift vs 13.0% base", "NCNS %"],
        ["0", "5,192", "10.4%", "0.80x", "2.9%"],
        ["1", "1,953", "11.6%", "0.89x", "3.7%"],
        ["2", "986", "12.8%", "0.98x", "3.5%"],
        ["3+", "1,582", "23.6%", "1.81x", "6.0%"],
    ], [2.0 * inch, 0.8 * inch, 1.4 * inch, 1.2 * inch, 0.7 * inch]))
    S.append(p("Workbook: Calc_Prediction (per-booking engine room: Booking_Calcs). The 3+ flag covers 16.3% of "
               "bookings and catches 29.5% of failures / 27.1% of NCNS.", cap))
    S.append(tbl([
        ["The last-minute claim correction", "Base", "Last-minute claims (n=1,306)", "Booked-ahead"],
        ["Shift-level “fail” (naive; includes failures that happened before the claim)", "17.5%", "31.3%", "15.4%"],
        ["Worker-level: this claimant late-cancels/NCNSes after booking", "13.0%", "12.0%", "13.2%"],
        ["Worker-level: any cancellation after booking", "24.4%", "12.0%", "26.3%"],
    ], [2.9 * inch, 0.7 * inch, 1.6 * inch, 1.0 * inch]))
    S.append(p(
        "Workbook: Calc_LastMin_Correction. The naive view brands last-minute claimers high-risk; the corrected view "
        "shows they are the most reliable segment. 354 of 1,306 last-minute claims (27%) were rescues of "
        "already-cancelled shifts; rescue claims ended with the shift worked 80.5% of the time.", cap))
    S.append(p(
        "First-timer wall: of 351 bookings ending in NCNS, 42.2% came from professionals with zero prior "
        "offenses and 69.8% from professionals with zero prior NCNS (workbook: Calc_Prediction). Detection "
        "lag: median NCNS event is logged 34.8h after shift start."))

    S.append(p("D. Assumptions register", h1))
    S.append(tbl([
        ["Assumption", "Why it is reasonable", "If wrong"],
        ["Verified timesheet = worked; unverified &amp; undeleted = empty",
         "Verified is the only ground truth for delivery; deletions are separated out",
         "Some “empty” shifts were worked unverified → prize shrinks proportionally; option ranking unchanged"],
        ["Final-event classification of multi-cancel shifts",
         "The last event determines the shift's ending state; avoids double counting",
         "Alternative (any-event) classification shifts totals ±10% but not the ordering of findings"],
        ["Booking log subset is representative of claim behavior",
         "Stated in the case brief; it is provided to observe HCP booking behavior",
         "Behavioral rates (12% vs 26%) could shift; refill economics (from shifts log) unaffected"],
        ["Cancel-log timestamps ≈ what our system knew in real time",
         "Events are logged at action time; it is our own event stream",
         "History-flag lift (1.8x) could attenuate; the first-timer wall (42–70%) is too large to flip"],
        ["Half the late→early refill gap is closable with speed",
         "Refill rises steeply with runway (25→45%); unmanaged rescues already work at 80.5%",
         "Full convergence is NOT assumed; targets and kill thresholds are pre-committed at day 90"],
        ["Cleveland generalizes across markets", "Stated in the case brief",
         "Rollout decision is gated on the Cleveland pilot readout regardless"],
    ], [1.9 * inch, 2.2 * inch, 2.2 * inch]))

    return S


def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#8a8a94"))
    canvas.drawString(0.85 * inch, 0.5 * inch, "Clipboard Health — Marketplace Reliability Case — D. Ezieshi")
    canvas.drawRightString(letter[0] - 0.85 * inch, 0.5 * inch, f"Page {doc.page}")
    canvas.restoreState()


def main():
    doc = BaseDocTemplate(str(OUT), pagesize=letter,
                          leftMargin=0.85 * inch, rightMargin=0.85 * inch,
                          topMargin=0.7 * inch, bottomMargin=0.75 * inch)
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="f")
    doc.addPageTemplates([
        PageTemplate(id="main", frames=[frame], onPage=on_page),
        PageTemplate(id="appendix", frames=[frame], onPage=on_page),
    ])
    doc.build(build())
    print("written", OUT)


if __name__ == "__main__":
    main()
