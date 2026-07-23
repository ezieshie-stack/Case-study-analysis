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
        "Late cancellations look like a discipline problem. A worker commits to a shift and then bails, so the "
        "instinct is to tighten the rules. We already tried that. The 2021 attendance policy helped a little "
        "and then stalled, and the data shows why. The reasons workers give for cancelling late are mostly "
        "things a penalty cannot touch: they are sick, a family member is in trouble, the car broke down. You "
        "cannot suspend someone out of a sick child. And every extra penalty we add pushes on the flexibility "
        "that brought them to Clipboard in the first place. Facilities and workers are both our customers, and "
        "a fix that wins one by losing the other is not a fix."))
    S.append(p(
        "The data points somewhere else. <b>What actually hurts a facility is not the cancellation. It is the "
        "shift that sits empty afterward.</b> When a worker cancels with a day or more of notice, the "
        "marketplace usually recovers on its own, and 67% of those shifts get claimed again and worked. Inside "
        "24 hours, recovery falls to 30%. After a no-call-no-show it falls to 14%. In this Cleveland window, "
        "2,634 shifts were abandoned by a worker and then never worked by anyone, and 89% of them trace back "
        "to late cancels and no-shows. That is roughly $730K of facility bookings gone in four months in one "
        "market, about $160K of our own revenue at a 22% take, spread across 55 of our 66 active facilities. "
        "Empty shifts are the thing we already know pushes facilities off the platform."))
    S.append(p(
        "So I want us to invest in one thing: winning the refill race. The moment a booked shift fails, "
        "whether it is a late cancel or a no-show, Clipboard should catch it within minutes and put it back in "
        "front of workers who have a track record of picking up same-day shifts. We already have proof this "
        "works. When a cancelled shift gets claimed by a same-day worker, it ends up worked 80% of the time, "
        "about the same as a shift that was never cancelled at all. Right now that only happens when we get "
        "lucky. I want to make it happen on purpose."))
    S.append(p(
        "My 90-day target is to cut cancel-driven empty shifts in Cleveland by about 104 a month. That is half "
        "the gap between how well we recover late cancels today and how well we already recover the early "
        "ones. In Cleveland alone it is worth roughly $85K a year in revenue we currently lose, before "
        "counting the churn we avoid, and the same mechanism carries over to every other market. It asks "
        "nothing punitive of workers."))

    # ------------------------------------------------------------------ 2
    S.append(p("2. Diagnosis: the harm is concentrated in the recovery gap", h1))
    S.append(p(
        "I used the shifts log as the base, 35,926 shifts after dropping 5,114 junk rows with a charge or "
        "duration of zero or less, and joined the booking and cancellation events onto it. The attached Excel "
        "models reproduce every table, definition, and cleaning step below, and Appendix D lists my "
        "assumptions."))
    S.append(p("2.1 Recovery depends almost entirely on notice", h2))
    S.append(tbl([
        ["Final cancel event on shift", "Shifts", "Refilled &amp; worked", "Died empty", "Deleted by facility"],
        ["Early cancel (≥24h notice)", "2,392", "67.4%", "11.8%", "20.8%"],
        ["Late cancel (&lt;24h notice)", "2,288", "30.2%", "60.1%", "9.7%"],
        ["No-call-no-show", "1,170", "13.8%", "83.4%", "2.7%"],
    ], [1.85 * inch, 0.7 * inch, 1.25 * inch, 0.9 * inch, 1.35 * inch]))
    S.append(p("Each shift counted once, classified by its final cancellation event. Source: models workbook, Calc_ShiftOutcomes tab.", cap))
    S.append(p(
        "Early cancellations mostly take care of themselves. Two-thirds get claimed again and worked with no "
        "help from us. The pain is concentrated where the notice runs out. Late cancels leave 1,376 empty "
        "shifts and no-shows another 976, and together those are 89% of every cancel-driven empty. (Another "
        "6,936 clean shifts were never claimed or cancelled at all and simply went unworked. That is a demand "
        "problem, not a reliability one, and nothing here rests on it.)"))
    S.append(p("2.2 Within the danger zone, speed is what separates recovery from loss", h2))
    S.append(tbl([
        ["Time remaining when late cancel lands", "Shifts", "Refilled &amp; worked", "Died empty"],
        ["12–24 hours", "316", "44.9%", "37.3%"],
        ["4–12 hours", "669", "33.5%", "57.2%"],
        ["Under 4 hours", "1,303", "24.9%", "67.2%"],
    ], [2.3 * inch, 0.7 * inch, 1.3 * inch, 1.0 * inch]))
    S.append(p("Source: models workbook, Calc_ShiftOutcomes tab.", cap))
    S.append(p(
        "Give the marketplace half a day and its odds of recovering a late cancel nearly double. The trouble "
        "is that the typical late cancel leaves only 3.1 hours before the shift starts, so every minute "
        "between the cancel and the re-offer costs us. No-shows are worse, and worse in a way we can actually "
        "fix. The median no-show is not even logged until about <b>35 hours after the shift was supposed to "
        "start</b>. We are not losing the race to refill those shifts. We are finding out there was a race a "
        "day too late."))

    # ------------------------------------------------------------------ 3
    S.append(p("3. Three tempting answers the data rules out", h1))
    S.append(p("3.1 Punish late cancellations harder", h2))
    S.append(p(
        "The policy is already live and the reasons are mostly genuine, but there is a nastier problem with "
        "escalating penalties. A worker who cannot cancel without getting punished can still just not show up. "
        "Every late cancel we push into a no-show takes a shift from a 60% chance of ending up empty to 83%, "
        "and takes the facility from a few hours of warning to none. Punishing harder risks manufacturing the "
        "exact outcome we most want to avoid."))
    S.append(p("3.2 Predict who will flake and intervene beforehand", h2))
    S.append(p(
        "I tested this one directly. For each of 9,713 bookings I counted how many late cancels and no-shows "
        "that worker already had before they made the booking, then checked whether they went on to fail it. "
        "There is a real signal, but only at the extreme. Workers with three or more prior offenses fail "
        "23.6% of their next bookings against a 13.0% base, about 1.8 times the rate, and I do use that flag. "
        "It just cannot be the whole plan, because <b>42% of no-shows come from workers with no prior offense "
        "of any kind, and 70% from workers who have never no-showed before</b>. Most of the worst events come "
        "from people no history-based system could have flagged. You can only prevent so much. There is no "
        "such ceiling on how much you can recover."))
    S.append(p("3.3 Add friction to last-minute bookings (and why I killed my own best finding)", h2))
    S.append(p(
        "My most eye-catching early finding was this: shifts claimed within 24 hours of the start fail 31.3% "
        "of the time, twice the 15.4% rate of shifts booked further out. The obvious move is to make "
        "last-minute claims harder or stickier. I almost proposed exactly that. Then I re-ran the test the "
        "right way, asking whether the person who made a given claim went on to cancel <i>that specific "
        "booking</i>, and the result flipped. Last-minute claimers only walk away from 12% of their claims. "
        "People who book days ahead walk away from 26%, because life gets in the way over those extra days. "
        "The 31.3% was backwards causation. About 27% of last-minute claims land on shifts somebody else had "
        "already cancelled, so the claim is the save. Those saved shifts get worked 80.5% of the time. Our "
        "same-day claimers are not the fragile ones. They are the people who bail us out, and putting friction "
        "in their way would break the very behavior this plan runs on. The full test is in Appendix C."))

    # ------------------------------------------------------------------ 4
    S.append(p("4. The solution: a rescue loop that makes recovery the default", h1))
    S.append(p(
        "Here is the whole thing in a sentence. <b>When a booked shift fails, Clipboard catches it within "
        "minutes and re-offers it right away to workers who reliably take same-day shifts, so it gets worked "
        "instead of going empty.</b> That pipeline has three parts. They are not three separate projects, and "
        "none of them does anything useful alone."))
    S.append(p(
        "<b>Detect the failure the moment it happens.</b> A late cancel already lands in our logs the instant "
        "it happens, but nothing operational fires off it today. A no-show we hear about a median of 35 hours "
        "late. So on the facility side I will add a one-tap \"worker hasn't shown up\" button that pops up "
        "automatically 15 minutes after the shift starts if the worker has not confirmed arrival in the app. "
        "The target: late cancels kick off the loop in under 5 minutes, and half of no-shows get reported "
        "within an hour of the start time, against roughly 2% today.", bullet))
    S.append(p(
        "<b>Re-offer it right away to the same-day pool.</b> The failed shift jumps to the top of an Urgent "
        "Shifts feed and fires targeted push notifications to workers who match on three things we know "
        "matter: the right license, a habit of claiming same-day work, and past shifts at that facility. This "
        "pool is not hypothetical. In the booking log, 4,259 different workers made a claim inside 24 hours, "
        "356 of them in Cleveland during this window, and those claims stick 88% of the time. To hit the "
        "target we need them to soak up something like three or four extra rescues a day across the market.", bullet))
    S.append(p(
        "<b>Start the race early for the shifts we can see coming.</b> The one predictive signal that held up "
        "under testing, three or more prior offenses, carries about 1.8 times the failure risk and covers "
        "roughly 16% of bookings. I use it to point the same pipeline earlier. A flagged booking gets a "
        "confirmation request 24 hours out. If the worker does not respond, nothing happens to them and the "
        "shift stays theirs. It just quietly shows up on the Urgent feed as \"backup wanted,\" so if the "
        "cancel does come, we start with hours of runway instead of minutes. This is a way to aim the rescue "
        "loop, not a second solution bolted on.", bullet))
    S.append(p(
        "Notice what the plan does not touch. No new penalty, no friction on booking, no new obligation on "
        "workers. Cancelling stays as flexible as it is now, and the current attendance policy stays exactly "
        "where it is. All I am doing is reacting to an event stream we already collect and steering open "
        "shifts toward workers who already behave this way. That is why it is cheap: about two engineer-months "
        "of work for the trigger, the feed, the push targeting, and the facility button, plus the ops time, "
        "which I will cover myself."))
    S.append(p(
        "I want to be straight about where this falls short, which is no-shows. They are 37% of the "
        "cancel-driven empties, they are mostly first offenses that no signal can catch in advance, and even "
        "when we detect one fast we can usually only salvage part of the shift. This plan recovers late "
        "cancels well and no-show hours only partly. I would rather say that plainly than go after the gap "
        "with penalties the data says will backfire."))

    # ------------------------------------------------------------------ 5
    S.append(p("5. Metrics: how we will know, by when", h1))
    S.append(p(
        "One caveat on the baseline before the targets. Late-cancel refill already climbed over this window, "
        "from 14% in October to 39% in January, as the market filled out. So I measure against January, the "
        "most recent month, rather than the kinder four-month average, and I track the <i>gap</i> to "
        "early-cancel refill so that any market-wide drift cancels out of the scorecard."))
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
    S.append(p("Guardrails, reviewed weekly, and I hold the authority to pause on any of them:", h2))
    S.append(p(
        "<b>Moral hazard.</b> If instant recovery makes cancelling feel free, the late-cancel rate per 100 "
        "booked shifts will creep up. I alert on a 10% relative rise and pause the rollout if it holds. "
        "<b>Supply cannibalization.</b> A rescue has to be a net-new fill, not a claim pulled off some other "
        "open shift, which is exactly why the headline metric is total empty shifts and not the refill rate on "
        "its own. <b>Worker experience.</b> Push notifications are capped per worker per day, confirmation "
        "requests only go to the flagged 16% of bookings, and the flag is never shown to facilities or tied to "
        "any penalty. <b>Facility trust.</b> The one-tap no-show report gets checked against timesheet data so "
        "it cannot be gamed."))

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
        "If the pipeline is clearly working, offers going out in minutes, but conversion still lags the "
        "target, then the problem is incentive rather than speed, and the next lever is a small same-day bonus "
        "paid for out of the roughly $75 we take on each rescued shift. I am leaving that spend out of this "
        "proposal on purpose. It is the obvious second iteration, and the data will tell us if we need it."))
    S.append(p(
        "At target this is about 104 more shifts worked every month in Cleveland, roughly $85K a year of "
        "revenue saved in one market, with the same machinery ready to drop into every other market. The part "
        "that matters most does not show up cleanly in that number. It is the facility that watches a failed "
        "shift get filled instead of living through an empty one. That is what keeps them with us."))

    # ------------------------------------------------------------------ appendix
    S.append(NextPageTemplate("appendix"))
    S.append(PageBreak())
    S.append(p("Appendix", title_st))
    S.append(p("Supporting analysis. The attached Excel models (Clipboard_Analysis_Models.xlsx) carry the shifts "
               "log and per-row analysis tabs, with each classification and join shown as its own column, "
               "feeding live-formula summary and sorted tabs. Every figure can be traced, filtered, and "
               "re-sorted in the workbook, with the tab named under each table.", sub_st))

    S.append(p("A. Data cleaning and definitions", h1))
    S.append(p(
        "<b>Universe:</b> 41,040 raw shifts; 5,114 rows removed for non-positive charge or duration (data "
        "errors); 35,926 clean shifts across 66 facilities, Oct 1 2021 – Jan 31 2022 starts; 91.5% CNA/LVN. "
        "The shifts table is the anchor; the booking (127,005 claims) and cancellation (78,073 events) logs "
        "are joined onto it. Both logs cover a wider date range than the shifts window, which is why raw join "
        "rates come in below 100%. <b>Worked</b> = Verified timesheet. <b>Facility-deleted</b> = Deleted flag "
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
        "Method: for each of the 9,713 bookings on the clean universe, I count a worker's offense history only "
        "up to the moment strictly before that booking, and the outcome I score is whether the same worker "
        "late-cancelled or no-showed that same booking after making it. Scoring it this way strips out two "
        "things that quietly corrupt a naive shift-level view: failures that happened before the claim (the "
        "rescue claims), and other workers' failures on the same shift."))
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
        ["Shift-level \"fail\" (naive; includes failures that happened before the claim)", "17.5%", "31.3%", "15.4%"],
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
         "Some \"empty\" shifts were worked unverified → prize shrinks proportionally; option ranking unchanged"],
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
