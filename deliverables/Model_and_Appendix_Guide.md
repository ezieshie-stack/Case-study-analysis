# How to use the Excel model and embed the exhibits

This is a short guide to the two things you asked about: which model figures to put in your appendix and how, why the Excel model matters, and how the analysis actually works so you can defend it in the interview.

---

## 1. What to embed, and where

Two exhibits are ready as images in `deliverables/figures/`:

- **`Fig_C1_refill_by_notice.png`** — the core diagnosis. Refill collapses as notice shrinks (67% early, 30% late, 14% no-show), and late cancels plus no-shows are 89% of the empty shifts. This is the single most important exhibit. It is the evidence for the whole recommendation.
- **`Fig_C2_lastminute_correction.png`** — the reverse-causality catch. The naive read says last-minute claimers are twice as risky; scored correctly they are the most reliable segment, and the claim is usually a rescue, not the risk.

Put both in **Appendix C** (the prediction-test section), right under the tables that are already there. C1 supports Section 2 (Diagnosis); C2 supports Section 3.3 (the finding you overturned).

### Embedding them in Word

1. In the Word doc (`Clipboard_Case_Proposal.docx`), click where you want the exhibit in the appendix.
2. Insert > Pictures > This Device, pick the PNG.
3. Right-click the image > Wrap Text > In Line with Text, and drag a corner to about 6.5 inches wide so it fits the margins.
4. Add a caption underneath in italics, for example:
   - Figure C1. Refill outcomes by notice, and where empty shifts come from. Source: Clipboard_Analysis_Models.xlsx, Calc_ShiftOutcomes tab.
   - Figure C2. The last-minute claim correction. Source: Clipboard_Analysis_Models.xlsx, Calc_LastMin_Correction tab.

A cleaner option, if you want it to be unmistakably your own work: open the workbook, go to those two tabs, and take your own screenshots. The figures I made reproduce the exact tab values, but a screenshot of your live model reads as more authentic and lets you crop it however you like. Either way, keep the "Source: tab name" line so a reader can click into the workbook and check the number.

### One rule for the submission

The PDF is the graded deliverable and it must stay at or under 6 pages of narrative. Exhibits go in the appendix, which can run past 6 pages. So embed these in the appendix, not the body.

---

## 2. Why the Excel model matters

The case rewards intellectual honesty and being "obvious to the reader," meaning someone can pick up your work and verify it. A deck of final numbers does not do that. A live model does. Three reasons this particular workbook earns its place:

1. **Every headline number is a formula, not a typed-in result.** The summary tabs are COUNTIFS, SUMIFS and AVERAGEIFS over the row-level data. An interviewer can click any cell and watch it trace back to the raw logs. Nothing is asserted that cannot be recomputed on the spot.
2. **It shows the working, not just the answer.** The analysis tabs carry one row per shift and per booking, with each classification and join as its own labeled column (worked, empty, final cancel class, prior-offense count, worker-level outcome). That is the difference between "trust me, it's 30%" and "here is every shift that rolls up into the 30%."
3. **It survives the one question that kills most of these cases.** The last-minute finding looked strong and was wrong. The model is what let me catch that, because I could re-score the same 9,713 bookings a different way and watch the number flip. If the analysis had only lived in a slide, the error would have shipped.

That last point is the real argument for the model. It is not decoration. It is the thing that caught my own mistake before an interviewer could.

---

## 3. The approach, walked through

The workbook is built in three layers, on purpose, so it opens fast and stays auditable.

**Layer 1, the raw log.** The shifts log is embedded whole, 41,040 rows, sorted and filterable. That lets a reader verify the cleaning step for themselves: 41,040 raw, drop 5,114 junk rows (charge or duration of zero or less), 35,926 clean. Nothing is hidden behind a pivot.

**Layer 2, the analysis tabs.** `Shift_Analysis` is one row per clean shift; `Booking_Analysis` is one row per booking that joins to the clean universe. These carry the derived columns. The rule for each column is written at the bottom of the tab. Two of those rules are the whole case:

- *Classify each shift once, by its final cancel event.* A shift can be claimed, cancelled, re-claimed and cancelled again, so if you count events you double-count. Collapsing to the final event is what makes the tables reconcile (6,960 events become 5,850 shifts).
- *Score outcomes at the worker level, point-in-time.* A booking only counts as a failure if that same worker late-cancelled or no-showed that same booking after making it, and a worker's offense history is counted only up to strictly before the booking. That removes two traps: failures that happened before a claim (which is what made last-minute claims look risky), and leakage from counting the future.

**Layer 3, the summary and sort tabs.** These are the live COUNTIFS/SUMIFS over Layer 2. `Calc_ShiftOutcomes` produces the refill-by-notice table (Figure C1). `Calc_Prediction` runs the point-in-time test. `Calc_LastMin_Correction` is the before-and-after that shows the reverse causality (Figure C2). `Sort_Facilities` and `Sort_Workers_NCNS` are the sorted concentration views (who feels the pain, and how concentrated the no-show tail is).

The logic that ties it together, in one line: recovery depends on notice, most of the damage is in the low-notice bucket, you cannot predict most of it in advance because the worst offenders are first-timers, so the leverage is in recovering faster rather than preventing. Every step of that sentence has a tab behind it.

### If someone asks how it was built

The workbook was generated from the raw logs by `analysis/build_models_workbook.py`, and every figure is independently re-derived in `analysis/clipboard_reliability_analysis.py` as a cross-check. The two agree. So the numbers are not just internally consistent, they were computed twice by two different methods and matched.
