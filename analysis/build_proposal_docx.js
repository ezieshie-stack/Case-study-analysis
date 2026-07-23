// Builds deliverables/Clipboard_Case_Proposal.docx, an editable Word version of
// the proposal, mirroring the PDF's humanized voice and structured to satisfy the
// case's stated grading characteristics (work backwards, anti-incremental,
// intellectual honesty, metrics-driven, ownership). US Letter.

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageBreak,
} = require("docx");

const OUT = path.join(__dirname, "..", "deliverables", "Clipboard_Case_Proposal.docx");

const ACCENT = "8a3033";
const INK = "1a1a2b";
const GREY = "55555f";
const HEADFILL = "f0eef2";
const GRID = "c9c9d1";
const BODY_FONT = "Calibri";

const DXA_PAGE = 12240 - 2 * 1080; // letter width minus 0.75in margins each side = content width

function run(text, opts = {}) {
  return new TextRun({ text, font: BODY_FONT, size: opts.size || 21, bold: !!opts.bold,
    italics: !!opts.italic, color: opts.color || INK });
}

function para(children, opts = {}) {
  const runs = Array.isArray(children) ? children : [run(children, opts)];
  return new Paragraph({
    children: runs,
    spacing: { after: opts.after == null ? 140 : opts.after, line: 264 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    indent: opts.indent ? { left: 300 } : undefined,
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 260, after: 110 },
    children: [new TextRun({ text, font: BODY_FONT, size: 26, bold: true, color: ACCENT })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 170, after: 70 },
    children: [new TextRun({ text, font: BODY_FONT, size: 22, bold: true, color: INK })],
  });
}
function caption(text) {
  return new Paragraph({
    spacing: { after: 150, before: 40 },
    children: [new TextRun({ text, font: BODY_FONT, size: 16, italics: true, color: GREY })],
  });
}
// bold lead-in + body, as one justified paragraph (the JTBD / guardrail style)
function leadPara(lead, rest, indent = false) {
  return new Paragraph({
    spacing: { after: 120, line: 264 },
    alignment: AlignmentType.JUSTIFIED,
    indent: indent ? { left: 300 } : undefined,
    children: [
      new TextRun({ text: lead, font: BODY_FONT, size: 21, bold: true, color: INK }),
      new TextRun({ text: " " + rest, font: BODY_FONT, size: 21, color: INK }),
    ],
  });
}

function cell(text, { bold = false, widthDxa, fill, align } = {}) {
  return new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    shading: fill ? { type: ShadingType.CLEAR, fill, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 90, right: 90 },
    children: [new Paragraph({
      alignment: align || AlignmentType.LEFT,
      spacing: { after: 0, line: 240 },
      children: [new TextRun({ text, font: BODY_FONT, size: 17, bold, color: INK })],
    })],
  });
}

function table(rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: GRID };
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    borders: { top: border, bottom: border, left: border, right: border,
      insideHorizontal: border, insideVertical: border },
    rows: rows.map((r, ri) =>
      new TableRow({
        tableHeader: ri === 0,
        children: r.map((c, ci) =>
          cell(c, { bold: ri === 0, widthDxa: widths[ci], fill: ri === 0 ? HEADFILL : undefined,
            align: ci === 0 ? AlignmentType.LEFT : AlignmentType.LEFT })),
      })),
  });
}

const doc = new Document({
  creator: "David Ezieshi",
  title: "Marketplace Reliability Case: Project Proposal",
  styles: { default: { document: { run: { font: BODY_FONT, size: 21, color: INK } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
      },
    },
    children: buildBody(),
  }],
});

function buildBody() {
  const W = DXA_PAGE;
  const S = [];

  // ---- title
  S.push(new Paragraph({
    spacing: { after: 40 },
    children: [new TextRun({ text: "Reducing the cost of late cancellations by winning the refill race",
      font: BODY_FONT, size: 34, bold: true, color: INK })],
  }));
  S.push(new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({
      text: "Project proposal for executive leadership   |   David Ezieshi   |   Operations Strategy   |   February 1, 2022   |   Data: Cleveland, shift starts Oct 1 2021 - Jan 31 2022",
      font: BODY_FONT, size: 19, color: GREY })],
  }));

  // ---- 1
  S.push(h1("1. Summary of the recommendation"));
  S.push(para(
    "Late cancellations look like a discipline problem. A worker commits to a shift and then bails, so the " +
    "instinct is to tighten the rules. We already tried that. The 2021 attendance policy helped a little and " +
    "then stalled, and the data shows why. The reasons workers give for cancelling late are mostly things a " +
    "penalty cannot touch: they are sick, a family member is in trouble, the car broke down. You cannot suspend " +
    "someone out of a sick child. And every extra penalty we add pushes on the flexibility that brought them to " +
    "Clipboard in the first place. Facilities and workers are both our customers, and a fix that wins one by " +
    "losing the other is not a fix."));
  S.push(new Paragraph({
    spacing: { after: 140, line: 264 }, alignment: AlignmentType.JUSTIFIED,
    children: [
      run("The data points somewhere else. "),
      run("What actually hurts a facility is not the cancellation. It is the shift that sits empty afterward. ", { bold: true }),
      run("When a worker cancels with a day or more of notice, the marketplace usually recovers on its own, and " +
        "67% of those shifts get claimed again and worked. Inside 24 hours, recovery falls to 30%. After a " +
        "no-call-no-show it falls to 14%. In this Cleveland window, 2,634 shifts were abandoned by a worker and " +
        "then never worked by anyone, and 89% of them trace back to late cancels and no-shows. That is roughly " +
        "$730K of facility bookings gone in four months in one market, about $160K of our own revenue at a 22% " +
        "take, spread across 55 of our 66 active facilities. Empty shifts are the thing we already know pushes " +
        "facilities off the platform."),
    ],
  }));
  S.push(para(
    "So I want us to invest in one thing: winning the refill race. The moment a booked shift fails, whether it is " +
    "a late cancel or a no-show, Clipboard should catch it within minutes and put it back in front of workers who " +
    "have a track record of picking up same-day shifts. We already have proof this works. When a cancelled shift " +
    "gets claimed by a same-day worker, it ends up worked 80% of the time, about the same as a shift that was " +
    "never cancelled at all. Right now that only happens when we get lucky. I want to make it happen on purpose."));
  S.push(para(
    "My 90-day target is to cut cancel-driven empty shifts in Cleveland by about 104 a month. That is half the " +
    "gap between how well we recover late cancels today and how well we already recover the early ones. In " +
    "Cleveland alone it is worth roughly $85K a year in revenue we currently lose, before counting the churn we " +
    "avoid, and the same mechanism carries over to every other market. It asks nothing punitive of workers."));

  // ---- 2
  S.push(h1("2. Diagnosis: the harm is concentrated in the recovery gap"));
  S.push(para(
    "I worked backwards from the outcome we want, which is a facility that gets the worker it was promised. So I " +
    "used the shifts log as the base, 35,926 shifts after dropping 5,114 junk rows with a charge or duration of " +
    "zero or less, and joined the booking and cancellation events onto it. The attached Excel models reproduce " +
    "every table, definition, and cleaning step below, and Appendix D lists my assumptions."));

  S.push(h2("2.1 Recovery depends almost entirely on notice"));
  S.push(table([
    ["Final cancel event on shift", "Shifts", "Refilled & worked", "Died empty", "Deleted by facility"],
    ["Early cancel (24h+ notice)", "2,392", "67.4%", "11.8%", "20.8%"],
    ["Late cancel (under 24h)", "2,288", "30.2%", "60.1%", "9.7%"],
    ["No-call-no-show", "1,170", "13.8%", "83.4%", "2.7%"],
  ], [Math.round(W * 0.30), Math.round(W * 0.12), Math.round(W * 0.22), Math.round(W * 0.16), W - Math.round(W * 0.30) - Math.round(W * 0.12) - Math.round(W * 0.22) - Math.round(W * 0.16)]));
  S.push(caption("Each shift counted once, classified by its final cancellation event. Source: models workbook, Calc_ShiftOutcomes tab."));
  S.push(para(
    "Early cancellations mostly take care of themselves. Two-thirds get claimed again and worked with no help " +
    "from us. The pain is concentrated where the notice runs out. Late cancels leave 1,376 empty shifts and " +
    "no-shows another 976, and together those are 89% of every cancel-driven empty. (Another 6,936 clean shifts " +
    "were never claimed or cancelled at all and simply went unworked. That is a demand problem, not a reliability " +
    "one, and nothing here rests on it.)"));

  S.push(h2("2.2 Within the danger zone, speed separates recovery from loss"));
  S.push(table([
    ["Time remaining when late cancel lands", "Shifts", "Refilled & worked", "Died empty"],
    ["12-24 hours", "316", "44.9%", "37.3%"],
    ["4-12 hours", "669", "33.5%", "57.2%"],
    ["Under 4 hours", "1,303", "24.9%", "67.2%"],
  ], [Math.round(W * 0.40), Math.round(W * 0.14), Math.round(W * 0.24), W - Math.round(W * 0.40) - Math.round(W * 0.14) - Math.round(W * 0.24)]));
  S.push(caption("Source: models workbook, Calc_ShiftOutcomes tab."));
  S.push(new Paragraph({
    spacing: { after: 140, line: 264 }, alignment: AlignmentType.JUSTIFIED,
    children: [
      run("Give the marketplace half a day and its odds of recovering a late cancel nearly double. The trouble is " +
        "that the typical late cancel leaves only 3.1 hours before the shift starts, so every minute between the " +
        "cancel and the re-offer costs us. No-shows are worse, and worse in a way we can actually fix. The median " +
        "no-show is not even logged until about "),
      run("35 hours after the shift was supposed to start", { bold: true }),
      run(". We are not losing the race to refill those shifts. We are finding out there was a race a day too late."),
    ],
  }));

  S.push(h2("2.3 What I am assuming"));
  S.push(para(
    "I lean on the assumptions the case gives me, and I state them so a reader can see where the argument would " +
    "bend if any turned out wrong. Both no-shows and late cancels correlate with facility churn, which is why I " +
    "treat the empty shift, not the cancel itself, as the thing to kill. The top cancel reasons are mostly " +
    "legitimate (sickness, emergencies, transport, facility issues), which is why I steer away from punishment. " +
    "Our take rate is 22%, which I use for every dollar figure. Cleveland dynamics generalize, which is why a " +
    "mechanism that works here is worth building once and reusing. The 2021 attendance policy already deactivates " +
    "repeat offenders and only moved the needle partway, and reminders already ship in the app, so both obvious " +
    "levers are spent. My own analytical assumptions (verified equals worked, final-event classification, the " +
    "booking log as a representative sample, and half the refill gap being closable) are listed in Appendix D " +
    "with what happens to the recommendation if each is wrong."));

  // ---- 3
  S.push(h1("3. Three answers I ruled out"));
  S.push(h2("3.1 Punish late cancellations harder"));
  S.push(para(
    "The policy is already live and the reasons are mostly genuine, but there is a nastier problem with " +
    "escalating penalties. A worker who cannot cancel without getting punished can still just not show up. Every " +
    "late cancel we push into a no-show takes a shift from a 60% chance of ending up empty to 83%, and takes the " +
    "facility from a few hours of warning to none. Punishing harder risks manufacturing the exact outcome we most " +
    "want to avoid."));
  S.push(h2("3.2 Predict who will flake and step in early"));
  S.push(new Paragraph({
    spacing: { after: 140, line: 264 }, alignment: AlignmentType.JUSTIFIED,
    children: [
      run("I tested this one directly. For each of 9,713 bookings I counted how many late cancels and no-shows that " +
        "worker already had before they made the booking, then checked whether they went on to fail it. There is a " +
        "real signal, but only at the extreme. Workers with three or more prior offenses fail 23.6% of their next " +
        "bookings against a 13.0% base, about 1.8 times the rate, and I do use that flag. It just cannot be the " +
        "whole plan, because "),
      run("42% of no-shows come from workers with no prior offense of any kind, and 70% from workers who have never " +
        "no-showed before", { bold: true }),
      run(". Most of the worst events come from people no history-based system could have flagged. You can only " +
        "prevent so much. There is no such ceiling on how much you can recover."),
    ],
  }));
  S.push(h2("3.3 Add friction to last-minute bookings (and why I killed my own best finding)"));
  S.push(new Paragraph({
    spacing: { after: 140, line: 264 }, alignment: AlignmentType.JUSTIFIED,
    children: [
      run("My most eye-catching early finding was this: shifts claimed within 24 hours of the start fail 31.3% of " +
        "the time, twice the 15.4% rate of shifts booked further out. The obvious move is to make last-minute " +
        "claims harder or stickier. I almost proposed exactly that. Then I re-ran the test the right way, asking " +
        "whether the person who made a given claim went on to cancel "),
      run("that specific booking", { italic: true }),
      run(", and the result flipped. Last-minute claimers only walk away from 12% of their claims. People who book " +
        "days ahead walk away from 26%, because life gets in the way over those extra days. The 31.3% was backwards " +
        "causation. About 27% of last-minute claims land on shifts somebody else had already cancelled, so the " +
        "claim is the save. Those saved shifts get worked 80.5% of the time. Our same-day claimers are not the " +
        "fragile ones. They are the people who bail us out, and putting friction in their way would break the very " +
        "behavior this plan runs on. The full test is in Appendix C."),
    ],
  }));

  // ---- 4
  S.push(h1("4. The solution: a rescue loop that makes recovery the default"));
  S.push(new Paragraph({
    spacing: { after: 130, line: 264 }, alignment: AlignmentType.JUSTIFIED,
    children: [
      run("Here is the whole thing in a sentence. "),
      run("When a booked shift fails, Clipboard catches it within minutes and re-offers it right away to workers " +
        "who reliably take same-day shifts, so it gets worked instead of going empty. ", { bold: true }),
      run("That pipeline has three parts. They are not three separate projects, and none of them does anything " +
        "useful alone."),
    ],
  }));
  S.push(leadPara("Detect the failure the moment it happens.",
    "A late cancel already lands in our logs the instant it happens, but nothing operational fires off it today. " +
    "A no-show we hear about a median of 35 hours late. So on the facility side I will add a one-tap \"worker " +
    "hasn't shown up\" button that pops up automatically 15 minutes after the shift starts if the worker has not " +
    "confirmed arrival in the app. The target: late cancels kick off the loop in under 5 minutes, and half of " +
    "no-shows get reported within an hour of the start time, up from 16.5% today.", true));
  S.push(leadPara("Re-offer it right away to the same-day pool.",
    "The failed shift jumps to the top of an Urgent Shifts feed and fires targeted push notifications to workers " +
    "who match on three things we know matter: the right license, a habit of claiming same-day work, and past " +
    "shifts at that facility. This pool is not hypothetical. In the booking log, 4,259 different workers made a " +
    "claim inside 24 hours, 356 of them in Cleveland during this window, and those claims stick 88% of the time. " +
    "To hit the target we need them to soak up something like three or four extra rescues a day across the market (about 104 a month).", true));
  S.push(leadPara("Start the race early for the shifts we can see coming.",
    "The one predictive signal that held up under testing, three or more prior offenses, carries about 1.8 times " +
    "the failure risk and covers roughly 16% of bookings. I use it to point the same pipeline earlier. A flagged " +
    "booking gets a confirmation request 24 hours out. If the worker does not respond, nothing happens to them " +
    "and the shift stays theirs. It just quietly shows up on the Urgent feed as \"backup wanted,\" so if the " +
    "cancel does come, we start with hours of runway instead of minutes. This is a way to aim the rescue loop, " +
    "not a second solution bolted on.", true));
  S.push(para(
    "Notice what the plan does not touch. No new penalty, no friction on booking, no new obligation on workers. " +
    "Cancelling stays as flexible as it is now, and the current attendance policy stays exactly where it is. All " +
    "I am doing is reacting to an event stream we already collect and steering open shifts toward workers who " +
    "already behave this way. That is why it is cheap: about two engineer-months of work for the trigger, the " +
    "feed, the push targeting, and the facility button, plus the ops time, which I will cover myself."));
  S.push(para(
    "I want to be straight about where this falls short, which is no-shows. They are 37% of the cancel-driven " +
    "empties, they are mostly first offenses that no signal can catch in advance, and even when we detect one " +
    "fast we can usually only salvage part of the shift. This plan recovers late cancels well and no-show hours " +
    "only partly. I would rather say that plainly than go after the gap with penalties the data says will backfire."));

  // ---- 5
  S.push(h1("5. Metrics: how we will know, by when"));
  S.push(new Paragraph({
    spacing: { after: 140, line: 264 }, alignment: AlignmentType.JUSTIFIED,
    children: [
      run("One caveat on the baseline before the targets. Late-cancel refill rose over this window on net, " +
        "from 14% in October to 39% in January (it slipped to 28% in December before recovering), as the market " +
        "filled out. So I measure against January, the most recent month, rather than the kinder four-month " +
        "average, and I track the "),
      run("gap", { italic: true }),
      run(" to early-cancel refill so that any market-wide drift cancels out of the scorecard."),
    ],
  }));
  S.push(table([
    ["Metric", "Baseline (Jan 2022)", "Day-90 target", "Success / failure line"],
    ["North star: cancel-driven empty shifts per week (Cleveland)", "~134", "≤110", "Success ≤115; failure >125 = kill or redesign"],
    ["Primary driver: late-cancel refill-to-worked, 30-day rolling", "39.3%", "≥53%", "Success ≥50%; failure <45%"],
    ["Gap to early-cancel refill benchmark (67.4%)", "28.1 pts", "≤14 pts", "Controls for organic drift"],
    ["Median time from failure event to re-offer", "no pipeline", "<5 min", "Leading indicator, weekly"],
    ["No-shows reported within 1h of shift start", "16.5%", "≥50%", "Leading indicator, weekly"],
  ], [Math.round(W * 0.34), Math.round(W * 0.16), Math.round(W * 0.15), W - Math.round(W * 0.34) - Math.round(W * 0.16) - Math.round(W * 0.15)]));
  S.push(h2("Guardrails, reviewed weekly, and I hold the authority to pause on any of them"));
  S.push(new Paragraph({
    spacing: { after: 140, line: 264 }, alignment: AlignmentType.JUSTIFIED,
    children: [
      run("Moral hazard. ", { bold: true }),
      run("If instant recovery makes cancelling feel free, the late-cancel rate per 100 booked shifts will creep " +
        "up. I alert on a 10% relative rise and pause the rollout if it holds. "),
      run("Supply cannibalization. ", { bold: true }),
      run("A rescue has to be a net-new fill, not a claim pulled off some other open shift, which is exactly why " +
        "the headline metric is total empty shifts and not the refill rate on its own. "),
      run("Worker experience. ", { bold: true }),
      run("Push notifications are capped per worker per day, confirmation requests only go to the flagged 16% of " +
        "bookings, and the flag is never shown to facilities or tied to any penalty. "),
      run("Facility trust. ", { bold: true }),
      run("The one-tap no-show report gets checked against timesheet data so it cannot be gamed."),
    ],
  }));

  // ---- 6
  S.push(h1("6. Execution: the first 90 days"));
  S.push(table([
    ["When", "What ships", "Owner"],
    ["Weeks 1-2 (Feb 1-14)", "Metrics instrumented from existing logs; baseline dashboard live; auto re-list trigger on late-cancel events, so the manual gap disappears", "Me + 1 backend eng"],
    ["Weeks 3-4", "Urgent Shifts feed and targeted push to same-day claimers, Cleveland pilot; start with the 10 facilities carrying half of late/NCNS empties", "Me + mobile eng (part-time)"],
    ["Weeks 5-6", "Facility one-tap no-show report and arrival-confirmation prompt", "Me + facility CS lead"],
    ["Weeks 7-8", "Confirmation requests 24h out on 3+ offense bookings; pre-warm flow live", "Me"],
    ["Weeks 9-13", "Weekly metric and guardrail reviews; iterate targeting; Day-90 readout with a scale-or-kill decision, pre-committed to the thresholds above", "Me, reporting to exec sponsor"],
  ], [Math.round(W * 0.20), W - Math.round(W * 0.20) - Math.round(W * 0.22), Math.round(W * 0.22)]));
  S.push(para(
    "If the pipeline is clearly working, offers going out in minutes, but conversion still lags the target, then " +
    "the problem is incentive rather than speed, and the next lever is a small same-day bonus paid for out of the " +
    "roughly $68 we take on each rescued shift. I am leaving that spend out of this proposal on purpose. It is the " +
    "obvious second iteration, and the data will tell us if we need it."));
  S.push(para(
    "At target this is about 104 more shifts worked every month in Cleveland, roughly $85K a year of revenue saved " +
    "in one market, with the same machinery ready to drop into every other market. The part that matters most " +
    "does not show up cleanly in that number. It is the facility that watches a failed shift get filled instead " +
    "of living through an empty one. That is what keeps them with us."));

  // ---- appendix
  S.push(new Paragraph({ children: [new PageBreak()] }));
  S.push(new Paragraph({
    spacing: { after: 40 },
    children: [new TextRun({ text: "Appendix", font: BODY_FONT, size: 30, bold: true, color: INK })],
  }));
  S.push(caption(
    "Supporting analysis. The attached Excel models (Clipboard_Analysis_Models.xlsx) carry the shifts log and " +
    "per-row analysis tabs, with each classification and join shown as its own column, feeding live-formula " +
    "summary and sorted tabs. Every figure can be traced, filtered, and re-sorted in the workbook, with the tab " +
    "named under each table."));

  S.push(h1("A. Data cleaning and definitions"));
  S.push(para(
    "Universe: 41,040 raw shifts; 5,114 rows removed for a charge or duration of zero or less; 35,926 clean " +
    "shifts across 66 facilities, October 2021 through January 2022 starts, 91.5% CNA or LVN. The shifts table is " +
    "the anchor, and the booking (127,005 claims) and cancellation (78,073 events) logs are joined onto it. Both " +
    "logs cover a wider date range than the shifts window, which is why raw join rates come in below 100%. Worked " +
    "means a verified timesheet. Facility-deleted means the Deleted flag (the facility cancelled it), and I keep " +
    "those out of the harm counts. Empty means neither worked nor facility-deleted. Early and late cancels are " +
    "worker cancels with 24h+ and under-24h lead. NCNS is the no-call-no-show action. A shift can be claimed, " +
    "cancelled, and re-claimed, so shifts with multiple cancel events are classified once, by the final event; " +
    "6,960 cancel events on the clean universe collapse to 5,850 distinct cancelled shifts, and every table sums " +
    "without double counting."));

  S.push(h1("B. Key volumes"));
  S.push(table([
    ["Quantity", "Value", "Workbook tab"],
    ["Cancel events (not shifts) on clean universe: early / late / NCNS", "3,347 / 2,436 / 1,177", "Calc_Notice"],
    ["Notice given, worker cancels: 72h+ / 24-72h / 4-24h / under 4h", "48.8% / 9.1% / 18.7% / 23.4%", "Calc_Notice"],
    ["Cancel-driven empty shifts (late / NCNS / early)", "1,376 / 976 / 282 = 2,634", "Calc_ShiftOutcomes"],
    ["Facility charge destroyed by late + NCNS empties (4 mo)", "$729,559 gross; ~$160K CBH take", "Calc_Prize"],
    ["Facilities hit by at least one late/NCNS empty; top-10 share", "55 of 66; 50.2%", "Sort_Facilities"],
    ["NCNS concentration: top 10% of NCNS workers", "30.9% of NCNS events", "Sort_Workers_NCNS"],
  ], [Math.round(W * 0.48), Math.round(W * 0.30), W - Math.round(W * 0.48) - Math.round(W * 0.30)]));

  S.push(h1("C. The prediction test (point-in-time, leakage-free)"));
  S.push(para(
    "Method: for each of the 9,713 bookings on the clean universe, I count a worker's offense history only up to " +
    "the moment strictly before that booking, and the outcome I score is whether the same worker late-cancelled " +
    "or no-showed that same booking after making it. Scoring it this way strips out two things that quietly " +
    "corrupt a naive shift-level view: failures that happened before the claim (the rescue claims), and other " +
    "workers' failures on the same shift."));
  S.push(table([
    ["Prior late/NCNS offenses at claim", "Bookings", "Fail % (after booking)", "Lift vs 13.0% base", "NCNS %"],
    ["0", "5,192", "10.4%", "0.80x", "2.9%"],
    ["1", "1,953", "11.6%", "0.89x", "3.7%"],
    ["2", "986", "12.8%", "0.98x", "3.5%"],
    ["3+", "1,582", "23.6%", "1.81x", "6.0%"],
  ], [Math.round(W * 0.30), Math.round(W * 0.14), Math.round(W * 0.22), Math.round(W * 0.20), W - Math.round(W * 0.30) - Math.round(W * 0.14) - Math.round(W * 0.22) - Math.round(W * 0.20)]));
  S.push(caption("Workbook: Calc_Prediction (per-booking engine room: Booking_Analysis). The 3+ flag covers 16.3% of bookings and catches 29.5% of failures / 27.1% of NCNS."));
  S.push(table([
    ["The last-minute claim correction", "Base", "Last-minute claims (n=1,306)", "Booked ahead"],
    ["Shift-level fail (naive; counts failures that predate the claim)", "17.5%", "31.3%", "15.4%"],
    ["Worker-level: this claimant late-cancels/NCNSes after booking", "13.0%", "12.0%", "13.2%"],
    ["Worker-level: any cancellation after booking", "24.4%", "12.0%", "26.3%"],
  ], [Math.round(W * 0.44), Math.round(W * 0.13), Math.round(W * 0.25), W - Math.round(W * 0.44) - Math.round(W * 0.13) - Math.round(W * 0.25)]));
  S.push(caption(
    "Workbook: Calc_LastMin_Correction. The naive view brands last-minute claimers high-risk; the corrected view " +
    "shows they are the most reliable segment. 354 of 1,306 last-minute claims (27%) were rescues of " +
    "already-cancelled shifts; rescue claims ended with the shift worked 80.5% of the time."));
  S.push(para(
    "First-timer wall: of 351 bookings ending in NCNS, 42.2% came from workers with zero prior offenses and 69.8% " +
    "from workers with zero prior NCNS (workbook: Calc_Prediction). Detection lag: the median NCNS event is logged " +
    "34.8 hours after shift start."));

  S.push(h1("D. Assumptions register"));
  S.push(table([
    ["Assumption", "Why it is reasonable", "If it is wrong"],
    ["Verified timesheet = worked; unverified and undeleted = empty", "Verified is the only ground truth for delivery; deletions are separated out", "Some empty shifts were worked unverified; the prize shrinks proportionally but the option ranking does not change"],
    ["Final-event classification of multi-cancel shifts", "The last event determines the shift's ending state and avoids double counting", "An any-event classification shifts totals by about 10% but not the ordering of the findings"],
    ["Booking-log subset represents claim behavior", "Stated in the case brief; it is provided to observe HCP booking behavior", "Behavioral rates (12% vs 26%) could shift; the refill economics from the shifts log are unaffected"],
    ["Cancel-log timestamps ≈ what the system knew in real time", "Events are logged at action time; it is our own event stream", "The history-flag lift (1.8x) could soften, but the first-timer wall (42-70%) is far too large to flip"],
    ["Half the late-to-early refill gap is closable with speed", "Refill rises steeply with runway (25% to 45%), and unmanaged rescues already work at 80.5%", "Full convergence is not assumed; targets and kill thresholds are pre-committed at day 90"],
    ["Cleveland generalizes across markets", "Stated in the case brief", "The rollout decision is gated on the Cleveland pilot readout regardless"],
  ], [Math.round(W * 0.30), Math.round(W * 0.35), W - Math.round(W * 0.30) - Math.round(W * 0.35)]));

  return S;
}

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("written", OUT, buf.length, "bytes");
});
