# Clipboard Health — Marketplace Reliability Case
## Analysis & Decision Journey

*Working document — the reasoning record behind the submission. Market: Cleveland. Period: Oct 2021 – Jan 2022 (shift start dates). All figures derived from the three provided logs (shifts, bookings, cancellations), cleaned to remove junk rows (zero/negative charge and duration). Canonical numbers are produced by `clipboard_reliability_analysis.py` and live in `deliverables/Clipboard_Analysis_Models.xlsx`; the submission itself is `deliverables/Clipboard_Case_Proposal.pdf`.*

---

## 1. The problem, in plain terms

Clipboard is a two-sided marketplace. Facilities post healthcare shifts; workers claim the ones they want; Clipboard takes ~22%. The problem: workers claim shifts and then bail — sometimes days ahead, sometimes hours ahead, and sometimes they simply never show up (a No-Call-No-Show). Late abandonment leaves a facility short-staffed for patient care with no time to recover, and repeated pain drives facilities off the platform.

**The constraint that shapes everything:** both sides are customers. Workers are on the app *for the flexibility*. Any fix that erodes that flexibility attacks Clipboard's own supply. So the solution must protect facility certainty **without** punishing the worker base.

**What's already been tried (and is therefore "spent"):**
- An attendance policy (deactivates repeat offenders) — moved the needle only partly.
- Pre-shift reminders — already in the app.

Both of the obvious answers (punish harder, remind more) are pre-empted. The real answer had to be found elsewhere.

---

## 2. What we're working with (data orientation)

- **41,040 posted shifts** (raw) across ~997 workers and 67 facilities; **35,926 clean** after removing 5,114 junk rows.
- Workforce is CNA/LVN-heavy (91.5% of clean shifts).
- The shifts table is the anchor universe; booking and cancel logs extend beyond its window, so they are joined *onto* the shifts table.
- Only 45.5% of clean shifts are verified as worked; 27.8% are deleted by the facility itself.
- 9,570 clean shifts ended **empty** (not worked, not facility-deleted). Of these, 2,634 followed a worker cancel/NCNS (this case's scope); 6,936 never had a cancel event at all — mostly never-claimed supply, a demand-fill problem outside the late-cancellation case.

---

## 3. The reasoning journey (blind → context → data)

**Blind instinct (no data):** add accountability — harsher penalties / reputation scores. *But the case already tried penalties and they underperformed, and the two intuitive levers are spent. The naive answer is the wrong one.*

**Context layer (before numbers):** This is healthcare (asymmetric stakes), 22% take (real revenue per shift), both sides are customers, and the reported cancel reasons are mostly **legitimate and unpreventable** (sickness, family emergencies, transport, facility issues). You cannot punish someone out of a sick child. This alone points away from prevention-by-punishment and toward **earlier warning + faster recovery**.

**Data layer:** confirmed and located the fragility, resized the prize, then — in the final verification pass — caught and corrected a reverse-causality error in our own "best" finding (§5.6).

---

## 4. Key findings from the data

### 4a. How much notice do workers give?
Most give plenty; a hard core give almost none.

| Notice window | % of worker cancels |
|---|---|
| <4h (basically no notice) | 23.4% |
| 4–24h | 18.7% |
| 24–72h | 9.1% |
| 72h+ | 48.8% |

Median notice ≈ 65 hours. It's a barbell: ~49% considerate, ~23% ambush.

### 4b. Cancellation mix — events vs shifts (reconciled)
Event level (6,960 cancel events on the clean universe): early 3,347 (48.1%), late 2,436 (35.0%), NCNS 1,177 (16.9%). A shift can be cancelled more than once (claim → cancel → re-claim → cancel), so events > shifts: the 6,960 events collapse to **5,850 distinct cancelled shifts**. All shift-level tables classify each shift **once, by its final cancel event** — this is what reconciles the earlier 3,347-vs-2,812 discrepancy between tabs.

### 4c. THE pivotal finding — refill depends entirely on notice
Final-event classification, each shift counted once:

| Type | Shifts | Refilled & worked | Died empty |
|---|---|---|---|
| Early cancel (≥24h) | 2,392 | **67.4%** | 11.8% |
| Late cancel (<24h) | 2,288 | **30.2%** | 60.1% |
| No-show | 1,170 | **13.8%** | 83.4% |

The system already recovers early cancels well. The bleeding is concentrated in **late cancels and no-shows that never get refilled**. Two sharpening facts from verification:
- **Runway gradient:** late cancels with 12–24h left refill at 44.9%; 4–12h → 33.5%; <4h → 24.9%. Median remaining runway on a late cancel: **3.1 hours**. Speed is the lever.
- **NCNS detection lag:** the median NCNS event is logged **~35 hours after shift start**. No-show recovery fails partly because we learn about the race after it's over.

### 4d. Where empty shifts actually come from (sizing the prize)

| Source (final event) | Empty shifts | Share of cancel-driven empty |
|---|---|---|
| Late cancels | 1,376 | **52.2%** |
| No-shows | 976 | **37.1%** |
| Early cancels | 282 | 10.7% |

**89% of the damage is late cancels + no-shows.** In dollars: the late+NCNS empties destroyed ~$730K of gross facility bookings (~$160K CBH take at 22%) in 4 months, in one market, touching 55 of 66 active facilities (top 10 facilities absorb half the pain).

**Lesson: rate ≠ volume.** The "striking" last-minute-claim finding (§5.6) sized out at ~116 empty shifts (~4%) even before it was overturned entirely. The boring-but-large finding wins.

---

## 5. The candidate solutions (BABOK "Define Change Strategy")

| Option | What it does | Verdict |
|---|---|---|
| **A — Booking friction** | Add commitment friction to last-minute claims | Killed twice over: tiny prize (~116 empties) AND the premise inverted under worker-level testing (§5.6). Last-minute claimers are the *most* reliable segment. **Dead.** |
| **B — Refill race** | Detect failures in minutes; re-offer instantly to proven same-day claimers | Biggest prize (~104 shifts/month at half-gap), near-zero worker risk, supply already exists and already behaves this way. **Winner.** |
| **C — No-show crackdown** | Attack no-shows via history/punishment | First-timer wall: 42% of NCNS come from workers with zero prior offenses, 70% with zero prior NCNS. Structurally unpredictable; punishment converts late cancels (60% empty) into no-shows (83% empty). Only the 3+ prior-offense flag survives, folded into B as a targeting input. |

### 5.5. The prediction test (the hinge)
Tested point-in-time (each worker's offense history counted strictly before the claim; no forward leakage), on 9,713 bookings joinable to the clean universe. First run defined "failure" at the **shift level** and appeared to show a strong last-minute-claim signal (31.3% vs 17.5% base). The final verification pass (§5.6) showed that definition was contaminated; the corrected, worker-level results are canonical:

- **Base rate:** 13.0% of bookings end with *that worker* late-cancelling or no-showing *after booking*. NCNS base: 3.6%.
- **History gradient:** 0 prior offenses → 10.4% | 1 → 11.6% | 2 → 12.8% | **3+ → 23.6% (1.81×, n=1,582)**. The flag covers 16.3% of bookings and catches 29.5% of failures.
- **First-timer wall (unchanged, the decider against prediction-led strategies):** 42.2% of NCNS bookings had zero prior offenses; 69.8% had zero prior NCNS.

### 5.6. The final verification pass — the reverse-causality catch
Re-deriving every number from the raw logs before submission exposed a flaw in our own "2× flake" finding:

| Definition | Base | Last-minute claims | Booked-ahead |
|---|---|---|---|
| Shift-level "fail" (naive) | 17.5% | **31.3%** | 15.4% |
| Worker-level: this claimant fails after booking | 13.0% | **12.0%** | 13.2% |
| Worker-level: any cancel after booking | 24.4% | **12.0%** | 26.3% |

The naive 31.3% counted failures that happened **before** the claim: **27% of last-minute claims are rescues of shifts someone else already cancelled** — the claim is the *cure*, not the *risk*. Rescued shifts get worked **80.5%** of the time (n=339), indistinguishable from never-cancelled shifts. Last-minute claimers abandon only 12% of their claims vs 26% for far-ahead bookers.

**Consequences:**
1. The last-minute-claim "risk signal" is **dead** — worse than dead for Option A: adding friction to last-minute claims would sabotage the marketplace's own recovery engine.
2. The same finding is **reborn as B's feasibility proof**: 4,259 distinct workers made sub-24h claims (356 in Cleveland in-window), their claims hold 88% of the time, and unmanaged rescues already work at 80.5%.
3. The proactive layer keeps exactly **one** tested signal: 3+ prior offenses (1.8×), used to start the refill race early (T-24h confirmation → pre-warm), never to punish.

---

## 6. Locked recommendation (final)

**Solution = win the refill race.** When a booked shift fails — late cancel or no-show — Clipboard detects it within minutes and re-offers it instantly to proven same-day claimers, so the shift is worked instead of dying empty. One pipeline, three jobs:

1. **Detect fast:** late cancels trigger at t=0 (event already logged; nothing fires today). No-shows: facility one-tap "hasn't arrived" report prompted 15 min after start (vs ~35h median detection today).
2. **Re-offer instantly:** Urgent Shifts feed + targeted push to workers with matching license, same-day claim history, and history at that facility. At target this pool absorbs ~3–4 extra rescues/day market-wide.
3. **Start early where possible:** 3+ prior-offense flag → T-24h confirmation request; no response quietly pre-lists the shift as "backup wanted." No penalty, no visibility to facilities, no release of the shift.

**The honest limit, stated in the proposal:** the NCNS residual. 37% of cancel-driven empties are no-shows; they're mostly first offenses no signal can flag, and rescue can usually save only part of the hours. The proposal recovers late cancels well, no-show hours partially, and says so.

**Prize (haircut, not ceiling):** close **half** the late→early refill gap from the honest baseline. Baseline = Jan 2022 exit rate (39.3% — refill improved organically 14%→39% over the window, so the 4-month average of 30% would flatter us). Half-gap to the 67.4% benchmark ≈ **104 shifts/month in Cleveland** (~$85K/yr take), before churn effects and before generalizing across markets. Full convergence (~208/month) is explicitly *not* claimed.

**Metrics (defined, dated, with kill thresholds):** north star = cancel-driven empty shifts/week (134 → ≤110 by day 90; failure >125 = kill/redesign); primary driver = late-cancel refill-to-worked 30-day rolling (39.3% → ≥53%; failure <45%); gap-to-benchmark tracked to net out drift; leading indicators = time-to-re-offer <5 min, ≥50% of NCNS reported within 1h. Guardrails: late-cancel rate per 100 bookings (moral-hazard alert at +10% relative), total-empty north star guards against supply cannibalization, push caps + non-punitive flag protect workers, one-tap reports audited against timesheets.

**Do NOT** simply ban late cancellations: with mostly-legitimate reasons, blocking late cancels converts them into no-shows — the *worst* outcome (83% empty vs 60%).

**One solution, not a bundle:** recovery *is* the solution; detection and the risk flag exist only to trigger the same re-offer pipeline sooner.

---

## 7. BABOK framing (write-up spine)

- **BACCM Need:** facilities churn because depended-on shifts die empty with no recovery time.
- **Change:** reduce *empty, unrecovered* shifts (not cancellations in general).
- **Value:** facility retention + revenue protection, without eroding worker flexibility.
- **Strategy Analysis:** current state ✅, future state ✅, risks ✅, change strategy → **refill race (reactive core) + one tested proactive signal (3+ prior offenses); last-minute-claim signal tested, overturned, and repurposed as feasibility evidence; no-show-history signal tested and excluded.**
- **Solution Evaluation:** metrics/dates/kill thresholds as in §6.

---

## 8. Open items — all closed

1. ~~Confirm early prediction works~~ — **ANSWERED (§5.5–5.6).** Mostly it doesn't (first-timer wall); one signal survives (3+ prior, 1.8×); the last-minute signal inverted under worker-level testing.
2. ~~Validate thresholds~~ — 3+ threshold documented as sample-fit; day-90 readout with pre-committed kill thresholds is the validation plan.
3. ~~Define metric targets and dates~~ — done (§6, proposal §5).
4. ~~Define guardrails~~ — done (§6, proposal §5).

---

## Appendix — method & honesty notes
- Junk rows (charge ≤ 0, time ≤ 0) removed before analysis.
- "Empty" defined as: not verified-worked AND not facility-deleted.
- Shifts table treated as anchor universe; logs joined onto it (explains apparent low raw join rates — logs cover a wider window).
- Multi-cancel shifts classified once, by final event (events 6,960 → shifts 5,850).
- Worker-level outcomes require the cancel event to postdate the booking (removes rescue-claim contamination).
- Point-in-time histories counted strictly before each booking timestamp (no leakage). Cancel-log timestamps are a directional proxy for real-time system knowledge; the first-timer wall (42–70%) is too large for proxy error to flip.
- External research was used only to sharpen mechanism thinking; per case instructions it is **not** cited in the proposal, and all conclusions derive from the provided data.
