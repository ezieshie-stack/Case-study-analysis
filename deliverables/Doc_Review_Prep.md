# Doc Review prep: everything behind the analysis

This is your study sheet. It walks the whole thing end to end so that when someone
pushes on any number, definition, or decision, you already know the answer and why
you made it. Read it until the logic feels like yours, because in the room the goal
is not to recite it but to reason from it.

The spine in one breath: facilities churn because shifts they were promised die
empty. Early cancels already refill fine; late cancels and no-shows are where the
empties come from, and they refill worse the less notice there is. You cannot
prevent most of them (the reasons are legitimate, and most no-shows are first
offenders), so the leverage is recovering faster, not punishing harder. That is
Option B, the refill race.

---

## Part 1 — The data you were given

Three files, all for the Cleveland market. Shift start dates run Oct 1 2021 to Jan 31 2022 (a 122-day window).

**Cleveland_shifts_logs — 41,040 rows, one per shift.** The anchor table. Columns:
Shift ID, Worker ID, Facility ID, Start, Agent Req (CNA/LVN/RN/…), End, Deleted,
Shift Type (AM/PM/NOC/custom), Created At (when the facility posted it), Verified
(TRUE if a signed timesheet confirms it was worked), Charge (per-hour rate),
Time (scheduled hours).

**Booking_logs — 127,005 rows, one per claim event.** Action ID, Created At (when
the claim happened), Shift ID, Action (always SHIFT_CLAIM), Worker ID, Facility ID,
Lead Time (hours between the claim and shift start). The brief says this log is only
a *subset* of the date range, so not every shift has a claim here; that is expected
and fine.

**Cancel_logs — 78,073 rows, one per cancellation event.** Action ID, Created At,
Shift ID, Action (WORKER_CANCEL or NO_CALL_NO_SHOW), Worker ID, Start, Facility ID,
Lead Time (hours between the cancel and shift start; negative means it was logged
*after* the shift already started).

**Things worth knowing cold:**
- No duplicate Action IDs in either log (I checked; 0 dupes). So dedup does not change the counts, but I dedup on Action ID anyway as a guard.
- "Deleted" means the *facility* cancelled the shift, not the worker. The footnote in the brief says so explicitly. This matters: a facility-deleted shift is not a worker reliability failure, so I exclude it from the harm.
- Lead Time is signed: `shift start − action time`, in hours. Positive = acted before start (notice given). Negative = acted after start (a no-show is logged hours *after* the shift began).

---

## Part 2 — Cleaning, and every definition

**Junk-row removal.** 41,040 → 35,926 clean shifts. I dropped 5,114 rows where Charge ≤ 0 or Time ≤ 0. A shift that bills nothing or lasts zero hours is a data artifact, not a real shift; leaving them in would distort revenue and rate math. If asked "why those two columns?" — they are the only two that can be non-positive and still parse as a number, and both being ≤ 0 is nonsensical for a real posted shift.

**The clean universe is my denominator for everything shift-level.** 35,926 shifts, ~66 facilities, CNA/LVN heavy (91.5% of clean shifts are CNA or LVN; the rest mostly RN).

**Definitions (memorize these — they are where an interviewer probes):**
- **Worked** = Verified is TRUE. A signed timesheet is the only ground truth that a shift was actually staffed. 45.5% of clean shifts are worked.
- **Facility-deleted** = the Deleted flag is set (=1). 27.8% of clean shifts. The facility pulled the shift; not a worker failure.
- **Empty** = clean AND not worked AND not facility-deleted. This is my harm unit. 9,570 clean shifts are empty.
- **Early cancel** = a WORKER_CANCEL with Lead Time ≥ 24h (a day or more of notice).
- **Late cancel** = a WORKER_CANCEL with Lead Time < 24h (under a day).
- **No-show (NCNS)** = the NO_CALL_NO_SHOW action.
- **Offense** = a late cancel or a no-show. Early cancels are not offenses (plenty of notice, marketplace recovers them).
- The 24-hour cut is the brief's own line ("late cancellations = cancels < 24 hours prior to shift start"), so I did not invent it.

**Why 9,570 empties but only 2,634 are "in scope."** Of the 9,570 empty shifts, only 2,634 followed a worker cancel or no-show. The other 6,936 were never claimed or cancelled at all — nobody ever picked them up. That is a *demand/supply-fill* problem (not enough workers wanted those shifts), which is a different case than late cancellations. I say this out loud in the proposal so nobody thinks I missed 6,936 empties; I scoped them out on purpose, and nothing in my recommendation depends on them.

---

## Part 3 — The join logic (and the "low join rate" question)

The shifts table is the anchor. I join booking and cancel events *onto* it by Shift ID.
The booking log (127k) and cancel log (78k) both cover a **wider date range** than the
shifts window, so most of their rows belong to shifts outside my window and do not
join. That is why the raw join looks low — it is not missing data, it is the logs
covering more calendar time than the shift file. Only **6,960 cancel events** land on
the clean universe. If asked "why is your join rate low?", that is the answer: the
logs are broader than the anchor, by design of how they were provided.

**One reconciliation you must be ready for.** 6,960 cancel *events* on the clean
universe, but only **5,850 distinct cancelled shifts**. Events > shifts because a
single shift can be claimed, cancelled, re-claimed, and cancelled again — multiple
cancel events on one Shift ID. So every shift-level table classifies each shift
**once, by its final cancel event** (the last one chronologically). That is the rule
that makes the tables sum without double-counting. Event-level counts (3,347 early /
2,436 late / 1,177 no-show = 6,960) live in the appendix and are labelled "events,
not shifts" precisely so nobody thinks the event table and the shift table contradict.

---

## Part 4 — How the thinking moved (the narrative they want to hear)

**Blind instinct, before any data:** add accountability — harsher penalties, a
reputation score. I rejected it fast, because the brief tells me the 2021 attendance
policy already deactivates repeat offenders and only moved the needle partway, and
reminders already ship in the app. The two obvious levers are spent.

**Context, before the numbers:** this is healthcare (asymmetric stakes — a missing
CNA is a patient-care problem), the take rate is 22% (real revenue per shift), both
sides are customers, and the brief says the top cancel reasons are sickness, family
emergencies, transport, and facility issues — mostly legitimate and unpreventable.
You cannot punish someone out of a sick child. That alone tilts the answer away from
prevention-by-punishment and toward earlier warning and faster recovery.

**Then the data,** in three moves: (1) locate the fragility, (2) size the prize —
which killed my most exciting finding, and (3) test the load-bearing assumption,
which caught a reverse-causality error in my own reframe. Parts 5–7 are those moves.

---

## Part 5 — Every finding, with the number and how it is computed

All of these are live formulas in `Clipboard_Analysis_Models.xlsx`, using named
ranges so they read in plain English. When someone asks "how did you get that?", the
answer is always "a COUNTIFS/SUMIFS over a labelled column — let me show you."

**5.1 Notice distribution (Calc_Notice).** Of worker cancels on the clean universe:
under 4h 23.4%, 4–24h 18.7%, 24–72h 9.1%, 72h+ 48.8%. Median notice ≈ 65 hours. Read:
it is a barbell — about half give plenty of notice (72h+), about a quarter give
almost none (<4h). `=COUNTIF(NoticeBucket,"<4h")` and so on.

**5.2 The pivotal table — refill by final cancel type (Calc_ShiftOutcomes).** Each
shift counted once:
| Final cancel type | Shifts | Refilled & worked | Died empty | Facility-deleted |
|---|---|---|---|---|
| Early (≥24h) | 2,392 | 67.4% | 11.8% | 20.8% |
| Late (<24h) | 2,288 | 30.2% | 60.1% | 9.7% |
| No-show | 1,170 | 13.8% | 83.4% | 2.7% |

Refill % is `=COUNTIFS(FinalClass,"late",ShiftWorked,1)` divided by
`=COUNTIF(FinalClass,"late")`. The story: recovery collapses as notice shrinks.
Early cancels basically fix themselves; the bleed is late cancels and no-shows.

**5.3 Runway inside the danger zone (Calc_ShiftOutcomes).** Among late cancels,
by hours of notice remaining: 12–24h → 316 shifts, 44.9% refilled; 4–12h → 669,
33.5%; under 4h → 1,303, 24.9%. Median late cancel leaves **3.1 hours**. Read: give
the marketplace half a day and its odds nearly double — speed is the lever. (Note the
bucket edges are half-open, ≥lo and <hi, on the *unrounded* lead time. This is why
the count is 669/1303, not 670/1302 — see Part 12, the rounding trap.)

**5.4 No-show detection lag (Calc_ShiftOutcomes).** Median no-show is logged **34.8
hours after** the shift was supposed to start. Only **16.5%** (194 of 1,177) are
logged within an hour of start; 33.3% within four hours. Read: for no-shows we lose
the refill race because we do not even know there was a race until the next day.

**5.5 Sizing the prize — where empties come from (Calc_ShiftOutcomes).** Of the 2,634
cancel-driven empties: late cancels 1,376 (52.2%), no-shows 976 (37.1%), early cancels
282 (10.7%). So **89% of the damage is late + no-show**. In money: those late+no-show
empties destroyed **$729,559** of gross facility charge over four months (late
$428,176 + no-show $301,383), which at the 22% take is **~$160,503** of Clipboard
revenue, across **55 of 66** facilities, with the top 10 facilities carrying half of
it. `=SUMIFS(GrossRev,FinalClass,"late",ShiftEmpty,1)`.

**5.6 The rate-vs-volume lesson.** My flashiest early finding (last-minute claims,
Part 6) had the highest failure *rate* but tiny *volume* — it sized to ~116 empty
shifts, about 4% of the total. The boring-but-large finding (late/no-show recovery)
is 20× the prize. This is the judgment move the interviewers look for: I resisted the
striking-but-small story in favour of the dull-but-big one.

**5.7 Concentration (Sort_Facilities, Sort_Workers_NCNS).** 55 of 66 facilities are
hit by at least one late/no-show empty; the top 10 absorb 50.2% of them. Among
no-show workers (429 of them), the top 10% (43 workers) cause 30.9% of no-show events.
That concentration is real and tempting — but see Part 6/7 for why it cannot carry a
proactive strategy.

---

## Part 6 — The reverse-causality catch (this is your best moment; know it cold)

**What I first found.** Shifts claimed within 24 hours of start fail 31.3% of the
time, versus 15.4% for shifts booked further out — last-minute claimers looked twice
as risky. The obvious move: add friction to last-minute claims. I almost proposed it.

**Why it is wrong.** That 31.3% is a *shift-level* number: it asks "did this shift end
in a late/no-show?" But a shift claimed last-minute is very often a shift that
*someone else already cancelled* — the late claim is the **rescue**, and the shift's
final-event label still reads "late/NCNS" from the earlier failure. So the naive test
was crediting the rescuer with the original failure. Reverse causality.

**How I proved it — re-score at the worker level.** Instead of "did this shift fail?",
I ask "did *this claimant* go on to late-cancel or no-show *this booking after they
made it*?" That is the only question that measures the claimant's own reliability:
| Definition of "failure" | Base | Last-minute claims | Booked-ahead |
|---|---|---|---|
| Shift-level (naive) | 17.5% | **31.3%** | 15.4% |
| Worker-level: this claimant fails after booking | 13.0% | **12.0%** | 13.2% |
| Worker-level: any cancel after booking | 24.4% | **12.0%** | 26.3% |

The finding **inverts**. Last-minute claimers abandon only 12% of their claims;
people who book days ahead abandon 26% (more time for life to intervene). Last-minute
claimers are the **most reliable** segment, not the least.

**The corroborating evidence (Calc_LastMin_Correction).** Of 1,306 last-minute claims,
354 (27.1%) were rescues of already-cancelled shifts; those rescue claims held (no
later cancel) 88% of the time and the shift ended up worked 80.5% of the time. So
the same behaviour I almost penalised is the recovery engine my recommendation runs on.

**Why this matters for the recommendation.** It flips Option A from "add friction to
last-minute claims" (which would sabotage rescues) into a proof of feasibility for
Option B (a proven same-day pool already exists and behaves exactly the way I need).

If they ask "how did you catch it?" — the honest answer: I re-derived every number
from the raw logs before submitting, and the last-minute finding only reproduced under
the shift-level definition. Changing the unit of analysis to the worker flipped it.

---

## Part 7 — The prediction test (methodology first, then results)

The question that had to be true for any *proactive* version of the plan: can I flag a
fragile shift before it fails, using only what is knowable at booking time?

**Method, and why it is leakage-free.** For each of the 9,713 bookings that join to
the clean universe, I count that worker's prior offenses (late cancels + no-shows)
**strictly before** the booking's timestamp — never after. Counting any offense at or
after the booking would be using the future to predict the past (leakage), which would
make any signal look artificially strong. The outcome I score is worker-level (did
*this* worker fail *this* booking after making it), same discipline as Part 6.

**Result 1 — base rates.** 13.0% of bookings end with that worker late-cancelling or
no-showing after booking; the no-show base is 3.6%.

**Result 2 — the offense-history gradient.**
| Prior offenses at claim | Bookings | Fail rate | Lift vs 13.0% base |
|---|---|---|---|
| 0 | 5,192 | 10.4% | 0.80× |
| 1 | 1,953 | 11.6% | 0.89× |
| 2 | 986 | 12.8% | 0.98× |
| 3+ | 1,582 | 23.6% | **1.81×** |

Signal only appears at the extreme: three or more prior offenses. That flag covers
16.3% of bookings and catches 29.5% of all failures (373 of them) and 27.1% of no-shows.

**Result 3 — the first-timer wall (the number that kills prediction-led plans).**
Of the bookings that end in a no-show, **42.2% came from workers with zero prior
offenses of any kind, and 69.8% from workers with zero prior no-shows**. Most of the
worst events come from people no history-based system could have flagged. You can only
prevent so much; there is no such ceiling on how much you can recover. That is the
data-grounded reason the plan is recovery-first, not prevention-first.

So proactive prediction is *partially* available: the 3+ flag is real and I use it, but
it can only ever touch a minority, and the no-show tail stays largely reactive.

---

## Part 8 — The three options and why B won

The brief demands one and only one solution. I scored three on prize, cost, and worker risk.

**A — Booking friction (add commitment cost to last-minute claims).** Dead twice
over: the prize is tiny (~116 empties, ~4%), and the premise inverted (Part 6) —
last-minute claimers are the reliable segment, so friction would break the recovery
engine. Rejected.

**B — The refill race (detect a failure fast, re-offer instantly to proven same-day
workers).** Biggest prize (~104 shifts/month at the half-gap target), near-zero worker
risk (nothing punitive, no new obligation), and the supply already exists and behaves
this way. **Winner.**

**C — No-show crackdown (punish/target no-shows).** Rejected as a standalone: 42–70%
of no-shows are first offenders, so you cannot predict most of them; and punishing
late cancels harder just converts them into no-shows, which are worse (83% empty vs
60%). The one salvageable piece — the 3+ prior-offense flag — folds into B as a
targeting input, not a separate initiative.

**Why not bundle.** The brief says bundling almost always falls short, and it is
right here: A adds friction at booking, B speeds recovery — two mechanisms, two
guardrail sets, two things to build. B is the highest-leverage single lever, so B is
the headline and the 3+ flag rides inside it as an input, not a co-solution.

---

## Part 9 — The recommendation in full (be able to whiteboard this)

**One sentence:** when a booked shift fails, Clipboard catches it within minutes and
re-offers it to workers who reliably take same-day shifts, so it gets worked instead
of going empty. One pipeline, three jobs:

**Detect fast.** A late cancel already hits our logs at t=0 but nothing operational
fires today. A no-show we hear about ~35h late. So I add a facility-side one-tap
"worker hasn't shown up" button, auto-prompted 15 minutes after start if the worker
hasn't confirmed arrival. Target: late cancels trigger the loop in under 5 minutes;
half of no-shows reported within an hour of start, up from 16.5% today.

**Re-offer instantly.** The failed shift jumps to the top of an Urgent Shifts feed and
fires targeted push to workers filtered on three things that matter: right license,
history of same-day claims, history at that facility. The pool is real — 4,259 distinct
workers made a sub-24h claim in the booking log (356 in Cleveland in-window), and those
claims stick 88% of the time.

**Start the race early where we can.** The one predictive signal that survived (3+
prior offenses, 1.8× risk, 16% of bookings) triggers a confirmation request 24h out.
No response does not punish the worker or release the shift — it quietly pre-lists it
as "backup wanted," so if the cancel comes we start with hours of runway instead of
minutes.

**What it deliberately does not do:** no new penalty, no booking friction, no new
worker obligation; the existing attendance policy is untouched. It consumes an event
stream we already have and routes demand to supply that already behaves this way,
which is why it is cheap (~two engineer-months plus ops time I staff myself).

**The honest limit:** no-shows are 37% of the empties, mostly first offenders, and
even detected fast we usually only salvage part of the shift. B recovers late cancels
well and no-show hours only partly, and I say so rather than pretend otherwise.

---

## Part 10 — Metrics, dates, thresholds, guardrails

**Baseline honesty first.** Late-cancel refill rose over the window on net, 14% (Oct)
→ 31% (Nov) → 28% (Dec, it dipped) → 39% (Jan). So I benchmark against **January**,
the most recent month, not the flattering four-month average of ~30%, and I track the
*gap* to early-cancel refill so market-wide drift nets out.

| Metric | Baseline (Jan 2022) | Day-90 target | Success / failure line |
|---|---|---|---|
| North star: cancel-driven empty shifts/week (Cleveland) | ~134 | ≤110 | success ≤115; failure >125 → kill/redesign |
| Late-cancel refill-to-worked, 30-day rolling | 39.3% | ≥53% | success ≥50%; failure <45% |
| Gap to early-cancel refill (67.4%) | 28.1 pts | ≤14 pts | controls for organic drift |
| Time from failure event to re-offer | no pipeline | <5 min | leading indicator |
| No-shows reported within 1h of start | 16.5% | ≥50% | leading indicator |

**The rescues-per-day sanity check.** Today the marketplace already recovers ~7 failed
shifts a day (853 late+no-show shifts worked over 122 days). The target adds ~3.4 more
a day (104/month ÷ 30.4), which is **~49% more — "roughly half again"**, routed
through a 4,259-worker pool. Be careful: 3.4 is the *increment* of recovered shifts;
the ~2.9/day figure in the workbook is a *different* unit (last-minute rescue *claims*,
one channel) and is flagged as a memo so the two are not confused.

**Guardrails (each with a trigger):** moral hazard (if instant recovery makes
cancelling feel free, late-cancel rate per 100 bookings rises — alert at +10% relative,
pause if sustained); supply cannibalization (rescues must be net-new fills, which is
why the north star is *total* empties, not the refill rate alone); worker experience
(push capped per worker/day; confirmation requests only to the flagged 16%; flag never
shown to facilities or tied to a penalty); facility trust (one-tap no-show reports
audited against timesheet data).

**Second iteration, pre-identified:** if the pipeline works but conversion lags, the
problem is incentive not speed, and the next lever is a small same-day bonus funded by
the ~$68 take on each rescued shift. Deliberately not bundled into this proposal.

---

## Part 11 — Assumptions, and what happens if each is wrong

| Assumption | Why reasonable | If wrong |
|---|---|---|
| Verified = worked; unverified & undeleted = empty | Verified timesheet is the only delivery ground truth | Some "empty" shifts were worked unverified → prize shrinks proportionally, ranking unchanged |
| Classify each multi-cancel shift by its final event | The last event is the shift's ending state; avoids double-counting | An any-event rule shifts totals ~10% but not the ordering of findings |
| Booking log is representative of claim behaviour | The brief says it is provided to observe HCP booking behaviour | Behavioural rates (12% vs 26%) could move; refill economics from the shifts log are unaffected |
| Cancel-log timestamps ≈ what the system knew in real time | They are our own event stream, logged at action time | The 1.8× history lift could soften; the 42–70% first-timer wall is far too large to flip |
| Half the late→early refill gap is closable with speed | Refill rises steeply with runway (25→45%); unmanaged rescues already work at 80.5% | Full convergence is NOT assumed; the day-90 kill thresholds are the real test |
| Cleveland generalises | The brief says to assume it | Rollout is gated on the Cleveland pilot readout regardless |

---

## Part 12 — Known weak points, own them before they are raised

- **The rounding trap (be ready, it is subtle).** The runway buckets are 316/669/1303.
  An earlier version of the workbook rounded the notice-hours column to 2 decimals,
  which pushed a 3.998h value onto 4.00 and across the bucket edge, giving 670/1302.
  I unrounded it so the workbook matches the proposal. If someone recomputes with
  rounding and gets 670/1302, that is the reason — the unrounded data has a clean gap
  at the boundary (values are 3.994, 3.998, then 4.002; nothing sits at exactly 4.00).
- **6,960 vs 6,962 cancel events.** The clean-universe count is 6,960 (the class counts
  sum to exactly that). A naive row-count of the Cancel_Analysis tab returns 6,962
  because it includes a blank spacer row and a footnote line with no data. Records =
  6,960.
- **No-show recovery is genuinely hard.** 83% die empty, mostly first offenders,
  detected late. The plan improves detection and recovers *some* hours; it does not
  solve no-shows. I state this rather than overclaim.
- **The $85K is Cleveland-only and pre-churn.** It is the direct recovered-revenue
  floor for one market at the half-gap target; the real value is churn avoided and the
  mechanism generalising, which I do not try to put a number on.
- **The 3+ threshold is fit on this sample.** It needs validation on held-out/future
  data before rollout; the day-90 readout is that validation, with pre-committed kill
  thresholds.

---

## Part 13 — Mock Doc Review: the hard questions and crisp answers

**"Why not just ban late cancellations?"** Because the reasons are mostly legitimate,
and a worker who cannot cancel can still no-show — which moves the shift from a 60%
chance of dying empty to 83%, and the facility from some warning to none. Banning
manufactures the worst outcome.

**"Your most interesting finding is the last-minute one — why isn't that the plan?"**
Two reasons. It is small — ~116 empties, 4% of the prize. And it is wrong the way I
first read it: re-scored at the worker level, last-minute claimers are the *most*
reliable segment, and 27% of those claims are rescues. Penalising them would break the
recovery engine.

**"Can't you just flag repeat offenders?"** Only partly. 42% of no-shows come from
workers with zero prior offenses and 70% from workers with no prior no-show, so most of
the worst events are unpredictable from history. The one flag that works, 3+ prior
offenses, catches under a third of failures. It is a targeting input, not a strategy.

**"Where does the $85K come from — walk me through it."** Jan baseline: 740 late-
cancelled shifts, 39.3% refilled. Early-cancel refill is 67.4%, so the gap is 28.1
points. Closing half of it recovers 740 × 28.1% × 0.5 ≈ 104 more shifts/month. At
~$309 gross per late shift × 22% take ≈ $68 each, that is ~$7,060/month ≈ $85K/year,
Cleveland only, before churn.

**"Is 104 a month even feasible for the worker pool?"** It is ~3.4 more recovered
shifts a day. Today the marketplace already recovers ~7 a day, so it is about half
again as much, spread over 4,259 workers who already claim same-day. It is a routing
problem, not a supply problem.

**"Why classify by the final cancel event and not the first?"** Because the shift's
ending state is what determines whether it was empty, and a shift can be cancelled,
re-claimed, and cancelled again. The final event avoids counting the same shift twice
and reconciles the event count (6,960) with the shift count (5,850).

**"Your refill improved on its own — how do you know your plan did anything?"** That is
exactly why I benchmark against January (the exit rate, 39.3%) not the four-month
average, and track the *gap* to early-cancel refill so any market-wide drift cancels
out of the scorecard. If the gap doesn't close, the plan isn't working, drift or not.

**"What's the single biggest risk to this plan?"** Moral hazard — that guaranteed fast
recovery makes cancelling feel consequence-free and the late-cancel rate creeps up. I
watch late-cancels per 100 bookings and pause at a sustained +10%. Second risk is
cannibalization (shuffling claims rather than net-new fills), which is why the north
star is total empty shifts, not the refill rate.

**"What would make you kill this?"** Day-90: north-star empties per week above 125, or
late-cancel refill below 45%. Pre-committed, in writing, before launch.

**"What did you get wrong during the analysis?"** The last-minute finding — I initially
read it as a risk signal and nearly built a solution on it. Sizing it and re-scoring it
at the worker level killed it. I would rather show that than pretend the first read was
right; it is the reason I trust the rest.

---

## Part 14 — Numbers cheat-sheet (one glance before you walk in)

- Universe: 41,040 raw → **35,926 clean** (dropped 5,114 junk). ~66 facilities. 91.5% CNA/LVN.
- Worked 45.5% · facility-deleted 27.8% · empty **9,570** (of which **2,634** cancel-driven, 6,936 never-claimed).
- Cancel events on clean universe: **6,960** → **5,850** distinct cancelled shifts.
- Shift-level (final event): early **2,392**, late **2,288**, no-show **1,170**.
- Refill: early **67.4%**, late **30.2%**, no-show **13.8%**. Empty: 11.8% / 60.1% / 83.4%.
- Empties by source: late **1,376** (52.2%), no-show **976** (37.1%), early 282 (10.7%).
- Runway (late): 12–24h **316**/44.9%, 4–12h **669**/33.5%, <4h **1,303**/24.9%. Median late notice **3.1h**.
- No-show detection: median **34.8h after** start; within 1h **16.5%** (194/1,177).
- Dollar harm: **$729,559** gross → **~$160.5K** take. Avg late shift $309; rescued shift $308 → **$68** take.
- Concentration: **55/66** facilities hit; top-10 = 50.2%. No-show top-10% workers = 30.9%.
- Prediction (9,713 bookings): base **13.0%**; 3+ prior offenses **23.6%** (1.81×), 16.3% of bookings, catches 29.5% of failures.
- First-timer wall: **42.2%** of no-shows zero prior offenses; **69.8%** zero prior no-shows.
- Last-minute correction: naive shift-level 31.3% vs worker-level **12.0%** (base 13.0%). 1,306 last-minute claims, **354** rescues (27%), held 88%, worked 80.5%.
- Target: Jan base **39.3%**, benchmark **67.4%**, gap **28.1pt**, half-gap **~104/month** ≈ **$85K/yr** Cleveland.
- Rescues/day: today ~**7** recovered, target **+3.4** (~49%, "half again").
