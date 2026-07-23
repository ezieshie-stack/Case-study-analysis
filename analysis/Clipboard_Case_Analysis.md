# Clipboard Health — Marketplace Reliability Case

## Working notes and how I got to the recommendation

These are my working notes, not the submission. Market is Cleveland, shift starts from October 2021 through January 2022. Every number comes from the three provided logs (shifts, bookings, cancellations), cleaned to drop junk rows with a charge or duration of zero or less. The canonical figures are produced by `clipboard_reliability_analysis.py` and live in `deliverables/Clipboard_Analysis_Models.xlsx`. The submission itself is `deliverables/Clipboard_Case_Proposal.pdf`.

---

## 1. The problem in plain terms

Clipboard runs a two-sided marketplace. Facilities post healthcare shifts, workers claim the ones they want, and Clipboard takes about 22%. The trouble is that workers claim shifts and then bail. Sometimes days ahead, sometimes hours ahead, and sometimes they just never turn up (a no-call-no-show). A late bail leaves a facility short-staffed for patient care with no time to fix it, and enough of that pain drives facilities off the platform.

The constraint that shapes everything: both sides are customers. Workers are on the app for the flexibility. Any fix that eats into that flexibility is really an attack on Clipboard's own supply. So whatever I propose has to give facilities more certainty without punishing the workers.

Two obvious answers are already spent. There is an attendance policy that deactivates repeat offenders, and it only moved things partway. And there are pre-shift reminders, which already exist in the app. Punish harder and remind more are both off the table, so the answer had to be somewhere else.

---

## 2. Getting oriented in the data

- 41,040 posted shifts in the raw file, about 997 workers and 67 facilities. After cleaning, 35,926 shifts.
- The workforce is mostly CNA and LVN, about 91.5% of clean shifts.
- The shifts table is the anchor. Booking and cancel logs run past its window, so I join them onto it.
- Only 45.5% of clean shifts are verified as worked. 27.8% are deleted by the facility itself.
- 9,570 clean shifts ended up empty (not worked, not facility-deleted). Of those, 2,634 followed a worker cancel or no-show, which is the part this case is about. The other 6,936 never had a cancel event at all, so they are mostly shifts nobody ever claimed. That is a demand problem, not a late-cancellation one.

---

## 3. How my thinking moved

Before I looked at anything, my gut said add accountability: harsher penalties, some kind of reputation score. But the case already tried penalties and they underperformed, and the two intuitive levers are used up. The naive answer was the wrong one, which is usually a good sign the interesting answer is elsewhere.

Adding context before the numbers pushed me further from prevention. This is healthcare, so the stakes are lopsided. The 22% take means there is real revenue riding on each shift. Both sides are customers. And the reasons workers give for last-minute cancels are mostly legitimate and unpreventable: sickness, family emergencies, transport, facility problems. You cannot punish someone out of a sick child. That alone points toward warning earlier and recovering faster rather than policing everyone.

Then the data. I confirmed where the fragility is, sized the prize, and in the final pass caught a reverse-causality error in my own best-looking finding. More on that in section 6.

---

## 4. What the data actually says

### 4a. How much notice do workers give

Most give plenty. A hard core give almost none.

| Notice window | Share of worker cancels |
|---|---|
| Under 4h (basically no notice) | 23.4% |
| 4–24h | 18.7% |
| 24–72h | 9.1% |
| 72h+ | 48.8% |

Median notice is about 65 hours. It is a barbell: roughly half are considerate, roughly a quarter are ambushes.

### 4b. Events versus shifts, reconciled

At the event level there are 6,960 cancel events on the clean universe: 3,347 early (48.1%), 2,436 late (35.0%), 1,177 no-show (16.9%). A shift can be cancelled more than once, since it can be claimed, cancelled, re-claimed, and cancelled again, so events outnumber shifts. Those 6,960 events collapse to 5,850 distinct cancelled shifts. Every shift-level table classifies a shift once, by its final cancel event. That is what reconciles the 3,347-versus-2,812 gap I had between two tabs early on.

### 4c. The finding the whole case turns on

Recovery depends almost entirely on notice. Each shift counted once:

| Type | Shifts | Refilled and worked | Died empty |
|---|---|---|---|
| Early cancel (24h+) | 2,392 | 67.4% | 11.8% |
| Late cancel (under 24h) | 2,288 | 30.2% | 60.1% |
| No-show | 1,170 | 13.8% | 83.4% |

The marketplace already recovers early cancels well. The bleeding is in late cancels and no-shows that never get refilled. Two things sharpened this when I dug in:

- Runway matters. Late cancels with 12 to 24 hours left refill at 44.9%. With 4 to 12 hours, 33.5%. Under 4 hours, 24.9%. The median late cancel leaves 3.1 hours. Speed is the lever.
- No-shows have a detection problem. The median no-show is not even logged until about 35 hours after the shift was supposed to start. Part of why no-show recovery fails is that we learn about the race after it is over.

### 4d. Where the empty shifts come from

| Source (final event) | Empty shifts | Share of cancel-driven empties |
|---|---|---|
| Late cancels | 1,376 | 52.2% |
| No-shows | 976 | 37.1% |
| Early cancels | 282 | 10.7% |

So 89% of the damage is late cancels plus no-shows. In dollars, the late and no-show empties destroyed about $730K of gross facility bookings, roughly $160K of Clipboard's take at 22%, over four months in one market, touching 55 of 66 active facilities. The top ten facilities absorb about half of it.

The lesson I kept coming back to is that rate is not volume. My flashiest early finding (the last-minute-claim angle) had the highest failure rate but small volume, and it sized out at about 116 empty shifts, roughly 4% of the total. Even before I found out it was wrong, it was the smallest prize on the board.

---

## 5. The options I weighed

| Option | What it does | Where it landed |
|---|---|---|
| A — Booking friction | Add commitment friction to last-minute claims | Dead twice over. Tiny prize (about 116 empties), and the premise inverted once I tested it properly (section 6). Last-minute claimers are actually the most reliable segment. |
| B — Refill race | Catch failures in minutes and re-offer them to proven same-day workers | The winner. Biggest prize (about 104 shifts a month at half the gap), near-zero worker risk, and the supply already exists and already behaves this way. |
| C — No-show crackdown | Go after no-shows with history or punishment | 42% of no-shows come from workers with no prior offense, 70% with no prior no-show. You cannot predict most of them from history. Punishment also converts late cancels (60% empty) into no-shows (83% empty). Only the 3+ prior-offense flag survives, and it folds into B as a targeting input. |

### 5.1. The test that decided it

I tested the load-bearing question point-in-time, counting each worker's offense history only up to strictly before the current claim so there is no leakage, on 9,713 bookings that join to the clean universe. My first cut defined failure at the shift level and seemed to show a strong last-minute signal, 31.3% against a 17.5% base. The final pass showed that definition was contaminated. The corrected, worker-level numbers are the ones I trust:

- Base rate: 13.0% of bookings end with that worker late-cancelling or no-showing after booking. No-show base is 3.6%.
- History gradient: 0 prior offenses gives 10.4%, one gives 11.6%, two gives 12.8%, and three-plus jumps to 23.6% (1.8 times base, on 1,582 bookings). The flag covers 16.3% of bookings and catches 29.5% of failures.
- The first-timer wall, which is the thing that kills prediction-led strategies: 42.2% of no-show bookings had zero prior offenses, and 69.8% had zero prior no-shows.

### 5.2. The reverse-causality catch

Re-deriving everything from the raw logs before submission is where I caught the flaw in my own "twice as likely to flake" finding:

| Definition | Base | Last-minute claims | Booked ahead |
|---|---|---|---|
| Shift-level fail (naive) | 17.5% | 31.3% | 15.4% |
| Worker-level: this claimant fails after booking | 13.0% | 12.0% | 13.2% |
| Worker-level: any cancel after booking | 24.4% | 12.0% | 26.3% |

The naive 31.3% was counting failures that happened before the claim. About 27% of last-minute claims are made on shifts someone else already cancelled, so the claim is the cure, not the cause. Those rescued shifts get worked 80.5% of the time, which is basically the same as a shift that was never cancelled. Last-minute claimers walk away from only 12% of their claims, against 26% for people who book far ahead.

Three things followed from that. The last-minute "risk signal" is dead, and worse than dead for option A, because adding friction to last-minute claims would sabotage the marketplace's own recovery engine. The same finding comes back as B's feasibility proof: 4,259 distinct workers made sub-24h claims, 356 in Cleveland in-window, their claims hold 88% of the time, and unmanaged rescues already work at 80.5%. And the proactive layer keeps exactly one tested signal, three-plus prior offenses at 1.8 times, used to start the refill race early, never to punish.

---

## 6. The recommendation I locked

Win the refill race. When a booked shift fails, whether a late cancel or a no-show, catch it within minutes and re-offer it to proven same-day workers so it gets worked instead of going empty. One pipeline, three jobs:

1. Detect fast. Late cancels already hit the logs at the moment they happen; nothing fires today. No-shows get a facility one-tap "hasn't shown up" prompt 15 minutes after start, against the roughly 35-hour median detection lag we have now.
2. Re-offer right away. An Urgent Shifts feed plus targeted push to workers who match on license, same-day claim history, and history at that facility. At target the pool needs to absorb three or four extra rescues a day across the market.
3. Start early where we can. The 3+ prior-offense flag triggers a confirmation request 24 hours out. No response does not punish the worker or release the shift; it quietly pre-lists the shift as backup wanted.

The honest limit is no-shows. They are 37% of cancel-driven empties, mostly first offenses no signal can flag, and rescue can usually save only part of the hours. The plan recovers late cancels well and no-show hours only partly, and I say so.

On the size of the prize, I lead with the haircut number, not the ceiling. I close half the gap between late-cancel refill and early-cancel refill, measured off the January baseline of 39.3% rather than the four-month average of 30% (refill improved on its own over the window, from 14% to 39%, so the average would flatter me). Half the gap to the 67.4% benchmark is about 104 shifts a month in Cleveland, roughly $85K a year in take, before churn and before other markets. Full convergence to 67.4% would be about 208 a month, and I do not claim it.

I am not banning late cancellations. With mostly legitimate reasons, blocking a late cancel just turns it into a no-show, which is the worst outcome (83% empty against 60%).

This is one solution, not a bundle. Recovery is the solution. Detection and the risk flag only exist to trigger the same re-offer pipeline sooner.

---

## 7. BABOK spine (for the write-up)

- Need: facilities churn because shifts they depend on die empty with no time to recover.
- Change: reduce empty, unrecovered shifts, not cancellations in general.
- Value: facility retention and revenue, without eroding worker flexibility.
- Strategy analysis: current state, future state, and risks all covered; strategy is the refill race as the reactive core plus one tested proactive signal (3+ prior offenses). The last-minute-claim signal was tested, overturned, and repurposed as feasibility evidence. The no-show-history signal was tested and excluded.
- Metrics, dates, and kill thresholds are in section 6 and the proposal.

---

## 8. Open items, now closed

1. Can we predict failures early enough to act? Answered in 5.1 and 5.2. Mostly not, because of the first-timer wall, but one signal works (3+ prior offenses), and the last-minute signal inverted under proper testing.
2. Validate the 3+ threshold. It is fit on this sample; the day-90 readout with pre-committed kill thresholds is the validation plan.
3. Metric targets and dates. Done, in section 6 and the proposal.
4. Guardrails. Done, in section 6 and the proposal.

---

## Method and honesty notes

- Junk rows (charge or time of zero or less) removed before analysis.
- Empty means not verified-worked and not facility-deleted.
- The shifts table is the anchor; the logs join onto it, which is why raw join rates look low (the logs cover a wider window).
- Multi-cancel shifts are classified once, by the final event (6,960 events collapse to 5,850 shifts).
- Worker-level outcomes require the cancel event to come after the booking, which removes the rescue-claim contamination.
- Point-in-time histories are counted strictly before each booking, so there is no leakage. Cancel-log timestamps are a directional proxy for what the system knew in real time; the first-timer wall (42 to 70%) is far too large for that proxy error to flip the conclusion.
- I used outside reading only to sharpen my thinking on mechanisms. Per the case instructions, none of it is cited in the proposal, and every conclusion comes from the provided data.
