"""
Builds deliverables/Clipboard_Analysis_Models.xlsx, the Excel models submitted
with the proposal.

Layered design (so the file opens and recalculates in seconds, yet every
headline number is a live, auditable formula):

  RAW LOGS      Shifts / Bookings / Cancels: the three logs verbatim, sorted,
                with AutoFilter so a reviewer can re-sort and slice them.

  ANALYSIS      Shift_Analysis (one row per clean shift) and Booking_Analysis
                (one row per booking joinable to the clean universe) carry the
                per-row derived columns: worked/empty flags, final cancel class,
                point-in-time prior-offense counts, worker-level outcome flags.
                These derivations are documented on each tab and reproduced
                independently by clipboard_reliability_analysis.py; they are
                written as values because a live point-in-time history count is a
                COUNTIFS over the 78k-row cancel log per booking (~10^9 cells) and
                is not tractable to recalc. Cancel_Analysis (anchor cancels) keeps
                its classification columns LIVE to show the rule in formula form.

  SUMMARY/SORT  Calc_* and Sort_* tabs produce every figure in the proposal with
                live COUNTIFS / SUMIFS / AVERAGEIFS over the analysis tabs (small
                ranges, instant recalc). Change an analysis-tab cell and the
                proposal numbers move.

Inputs (not committed): data/Cleveland_shifts_logs.xlsx, data/Booking_logs.xlsx,
data/Cancel_logs.xlsx.  Recalculate once after building.
"""

import os
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Which raw logs to embed as full data tabs. The analysis tabs already carry
# every source column for the rows in scope, so the two large logs are optional
# weight; embedding all three inflates the file past what a headless recalc can
# reopen. "shifts" (default) keeps the compact anchor log and drops the 127k-row
# Bookings and 78k-row Cancels raw dumps (the provider already has those files).
RAW_TABS = os.environ.get("RAW_TABS", "shifts")  # "shifts" | "all" | "none"

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "deliverables" / "Clipboard_Analysis_Models.xlsx"

ARIAL = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
TITLE = Font(name="Arial", size=12, bold=True)
BLUE = Font(name="Arial", size=10, color="0000FF")
NOTE = Font(name="Arial", size=9, italic=True, color="555555")
HFILL = PatternFill("solid", fgColor="DDD9E4")
PCT = "0.0%"
MONEY = "$#,##0"
DATEF = "yyyy-mm-dd hh:mm"


def header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = BOLD
        cell.fill = HFILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def dump_raw(ws, df, date_cols, widths):
    ws.append(list(df.columns))
    header(ws, 1, len(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))
    for col in date_cols:
        idx = list(df.columns).index(col) + 1
        letter = get_column_letter(idx)
        for r in range(2, len(df) + 2):
            ws[f"{letter}{r}"].number_format = DATEF
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w


def dump_values(ws, df, date_cols, pct_cols, money_cols, widths, note=None):
    ws.append(list(df.columns))
    header(ws, 1, len(df.columns))
    for row in df.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
    cols = list(df.columns)
    for name, fmt in [(c, DATEF) for c in date_cols] + [(c, PCT) for c in pct_cols] + [(c, MONEY) for c in money_cols]:
        if name in cols:
            letter = get_column_letter(cols.index(name) + 1)
            for r in range(2, len(df) + 2):
                ws[f"{letter}{r}"].number_format = fmt
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    if note:
        rr = len(df) + 3
        ws.cell(row=rr, column=1, value=note).font = NOTE


def block(ws, start, title, headers, rows, pct=(), money=(), mult=(), widths=None, note=None):
    r = start
    ws.cell(row=r, column=1, value=title).font = TITLE
    r += 1
    for j, h in enumerate(headers, start=1):
        ws.cell(row=r, column=j, value=h)
    header(ws, r, len(headers))
    r += 1
    for row in rows:
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = ARIAL
            if j in pct:
                c.number_format = PCT
            if j in money:
                c.number_format = MONEY
            if j in mult:
                c.number_format = "0.00x"
        r += 1
    if widths:
        for j, w in enumerate(widths, start=1):
            cur = ws.column_dimensions[get_column_letter(j)].width
            if cur is None or cur < w:
                ws.column_dimensions[get_column_letter(j)].width = w
    if note:
        ws.cell(row=r, column=1, value=note).font = NOTE
        r += 1
    return r + 1


def classify(action, lead):
    if action == "NO_CALL_NO_SHOW":
        return "NCNS"
    return "early" if lead >= 24 else "late"


def main():
    shifts = pd.read_excel(DATA / "Cleveland_shifts_logs.xlsx").sort_values("Shift ID").reset_index(drop=True)
    book = pd.read_excel(DATA / "Booking_logs.xlsx").drop_duplicates(subset="Action ID").sort_values("Created At").reset_index(drop=True)
    cancel = pd.read_excel(DATA / "Cancel_logs.xlsx").drop_duplicates(subset="Action ID").sort_values(["Shift ID", "Created At"]).reset_index(drop=True)

    cancel["Class"] = [classify(a, l) for a, l in zip(cancel["Action"], cancel["Lead Time"])]
    cancel["IsOffense"] = (cancel["Class"] != "early").astype(int)

    clean = shifts[(shifts["Charge"] > 0) & (shifts["Time"] > 0)].copy()
    clean_ids = set(clean["Shift ID"])
    clean["Worked"] = (clean["Verified"].fillna(False).astype(bool)).astype(int)
    clean["FacDeleted"] = (clean["Deleted"] == 1.0).astype(int)
    clean["Empty"] = ((clean["Worked"] == 0) & (clean["FacDeleted"] == 0)).astype(int)
    clean["GrossRev"] = (clean["Charge"] * clean["Time"]).round(2)

    c_anchor = cancel[cancel["Shift ID"].isin(clean_ids)].copy()
    last_ev = c_anchor.sort_values("Created At").groupby("Shift ID").last()
    clean["FinalCancelClass"] = clean["Shift ID"].map(last_ev["Class"]).fillna("none")
    lead_map = last_ev["Lead Time"].to_dict()
    # Do NOT round: rounding to 2 dp pushes a ~3.995h value onto 4.00 and moves it
    # across the runway-bucket boundary, so the workbook would disagree with the
    # unrounded figures in the proposal by one shift.
    clean["FinalCancelLead"] = clean["Shift ID"].map(lead_map)
    clean["StartMonth"] = clean["Start"].dt.strftime("%Y-%m")

    shift_an = clean[["Shift ID", "Worker ID", "Facility ID", "Start", "Agent Req", "Deleted",
                      "Verified", "Charge", "Time", "Worked", "FacDeleted", "Empty", "GrossRev",
                      "FinalCancelClass", "FinalCancelLead", "StartMonth"]].reset_index(drop=True)

    # ----- Booking_Analysis (point-in-time, worker-level) -----
    b = book[book["Shift ID"].isin(clean_ids)].copy().reset_index(drop=True)
    b["LastMinute"] = (b["Lead Time"] < 24).astype(int)

    off = cancel[cancel["IsOffense"] == 1][["Worker ID", "Created At", "Class"]].sort_values("Created At")
    off_by_w = {w: g["Created At"].values for w, g in off.groupby("Worker ID")}
    ncns_by_w = {w: g["Created At"].values for w, g in off[off["Class"] == "NCNS"].groupby("Worker ID")}

    def prior(d, w, t):
        a = d.get(w)
        return 0 if a is None else int(np.searchsorted(a, t, side="left"))

    b["PriorOffenses"] = [prior(off_by_w, w, t) for w, t in zip(b["Worker ID"], b["Created At"].values)]
    b["PriorNCNS"] = [prior(ncns_by_w, w, t) for w, t in zip(b["Worker ID"], b["Created At"].values)]

    c2 = cancel.merge(b[["Worker ID", "Shift ID", "Action ID", "Created At"]].rename(
        columns={"Action ID": "bid", "Created At": "bt"}), on=["Worker ID", "Shift ID"])
    after = c2[c2["Created At"] >= c2["bt"]]
    b["FailAfter"] = b["Action ID"].isin(set(after.loc[after["Class"].isin(["late", "NCNS"]), "bid"])).astype(int)
    b["NCNSAfter"] = b["Action ID"].isin(set(after.loc[after["Class"] == "NCNS", "bid"])).astype(int)
    b["AnyCancelAfter"] = b["Action ID"].isin(set(after["bid"])).astype(int)
    first_cx = c_anchor.groupby("Shift ID")["Created At"].min()
    b = b.merge(first_cx.rename("fc"), left_on="Shift ID", right_index=True, how="left")
    b["RescueClaim"] = (b["fc"].notna() & (b["Created At"] > b["fc"])).astype(int)
    b["ShiftWorked"] = b["Shift ID"].map(clean.set_index("Shift ID")["Worked"]).fillna(0).astype(int)
    b["ShiftFinalClass"] = b["Shift ID"].map(clean.set_index("Shift ID")["FinalCancelClass"]).fillna("none")
    b["ShiftFailedNaive"] = b["ShiftFinalClass"].isin(["late", "NCNS"]).astype(int)

    book_an = b[["Action ID", "Created At", "Shift ID", "Worker ID", "Lead Time", "LastMinute",
                 "PriorOffenses", "PriorNCNS", "FailAfter", "NCNSAfter", "AnyCancelAfter",
                 "RescueClaim", "ShiftWorked", "ShiftFinalClass", "ShiftFailedNaive"]].reset_index(drop=True)

    NSA = len(shift_an) + 1
    NBA = len(book_an) + 1
    NCA = len(c_anchor) + 1

    wb = Workbook()

    # ------------------------------------------------------------- ReadMe
    rm = wb.active
    rm.title = "ReadMe"
    L = [
        ("Clipboard Health, Marketplace Reliability Case: Excel models", TITLE),
        ("Market: Cleveland. Shift starts Oct 1 2021 to Jan 31 2022. Author: David Ezieshi.", ARIAL),
        ("", ARIAL),
        ("Three layers", BOLD),
        ("  RAW LOG (Shifts): the provided shifts log verbatim, sorted, with AutoFilter on (embeds all 41,040", ARIAL),
        ("    rows so the junk-row cleaning, 41,040 down to 35,926, is auditable). The Booking_logs and Cancel_logs", ARIAL),
        ("    are the files you provided; their in-scope rows (with every source column) are carried on the", ARIAL),
        ("    Booking_Analysis and Cancel_Analysis tabs, so nothing needed for the analysis is missing here.", ARIAL),
        ("  ANALYSIS: Shift_Analysis (one row per clean shift) and Booking_Analysis (one row per booking", ARIAL),
        ("    on the clean universe) hold the per-row derived columns. Each derivation rule is stated at the", ARIAL),
        ("    bottom of its tab and is reproduced independently by analysis/clipboard_reliability_analysis.py.", ARIAL),
        ("    Cancel_Analysis keeps its classification columns as LIVE formulas so you can see the rule.", ARIAL),
        ("  SUMMARY / SORT: every Calc_* and Sort_* figure is a LIVE COUNTIFS / SUMIFS / AVERAGEIFS over the", ARIAL),
        ("    analysis tabs. Edit an analysis cell and the proposal numbers recalculate.", ARIAL),
        ("", ARIAL),
        ("Named ranges (so the Calc_ formulas read in plain English)", BOLD),
        ("  Every column used in a summary has a name, so a formula reads =COUNTIFS(FinalClass,\"late\",ShiftWorked,1)", ARIAL),
        ("  instead of =COUNTIFS(Shift_Analysis!$N$2:$N$35927, ...). To see or edit them: Formulas > Name Manager.", ARIAL),
        ("  Shift rows:   FinalClass (early/late/NCNS), ShiftWorked, FacDeleted, ShiftEmpty, GrossRev, FinalLead", ARIAL),
        ("                (hours notice at final cancel), StartMonth, Facility.", ARIAL),
        ("  Booking rows: LastMinute (claimed <24h out), PriorOffenses, PriorNCNS, DidFail (this worker failed", ARIAL),
        ("                this booking), DidNoShow, DidCancelAfter, IsRescueClaim, BkShiftWorked, ShiftFailedNaive.", ARIAL),
        ("  Cancel rows:  CxClass (early/late/NCNS), NoticeBucket, CxWorker, CxLead (hours; negative = after start).", ARIAL),
        ("  Read a formula out loud: COUNTIFS just means 'count the rows where every condition is true'; SUMIFS", ARIAL),
        ("  means 'add up one column for the rows where every condition is true'; a lone /B3 divides by the total.", ARIAL),
        ("", ARIAL),
        ("Why the per-row history columns are values, not formulas", BOLD),
        ("  A live point-in-time prior-offense count is a COUNTIFS over the 78,000-row cancel log for every", ARIAL),
        ("  one of 9,713 bookings (about a billion cell tests), which makes the file take many minutes to open.", ARIAL),
        ("  They are computed once by the documented rule; the summary tables that use them stay live and instant.", ARIAL),
        ("", ARIAL),
        ("Load-bearing sort orders", BOLD),
        ("  Shifts: by Shift ID.  Cancels: by Shift ID then Created At (so the last row per shift is its final", ARIAL),
        ("  event).  Bookings: by Created At.", ARIAL),
        ("", ARIAL),
        ("Definitions", BOLD),
        ("  Clean universe: Charge > 0 AND Time > 0 (5,114 junk rows excluded).", ARIAL),
        ("  Worked = Verified TRUE.  FacDeleted = Deleted flag.  Empty = clean AND not worked AND not deleted.", ARIAL),
        ("  Early cancel >= 24h notice; Late < 24h; NCNS = no-call-no-show. Offense = late or NCNS.", ARIAL),
        ("  Final cancel class: each shift classified once, by its last cancel event.", ARIAL),
        ("  Worker-level outcome: a booking fails only if THAT worker late-cancels/NCNSes THAT shift at or", ARIAL),
        ("  after the booking time (removes reverse-causality). Prior offenses counted strictly before booking.", ARIAL),
        ("", ARIAL),
        ("Tab guide", BOLD),
        ("  Shifts ................................ raw shifts log (sorted, filterable)", ARIAL),
        ("  Cancel_Analysis ....................... anchor cancels; LIVE classification formulas", ARIAL),
        ("  Shift_Analysis ........................ per-shift derived columns + rules", ARIAL),
        ("  Booking_Analysis ...................... per-booking derived columns + rules (prediction engine room)", ARIAL),
        ("  Calc_ShiftOutcomes .................... refill by type, empty sources, runway, monthly trend", ARIAL),
        ("  Calc_Notice ........................... cancel-event mix and notice-given distribution", ARIAL),
        ("  Calc_Prediction ....................... base rates, offense-history gradient, first-timer wall", ARIAL),
        ("  Calc_LastMin_Correction ............... shift- vs worker-level last-minute test + rescue evidence", ARIAL),
        ("  Calc_Prize ............................ $ sizing and the 90-day target arithmetic", ARIAL),
        ("  Sort_Facilities ....................... facilities ranked by late/NCNS empties, cumulative share", ARIAL),
        ("  Sort_Workers_NCNS ..................... workers ranked by NCNS events, top-10% concentration", ARIAL),
    ]
    for i, (t, f) in enumerate(L, start=1):
        rm.cell(row=i, column=1, value=t).font = f
    rm.column_dimensions["A"].width = 108

    # ------------------------------------------------------------- raw logs
    if RAW_TABS in ("shifts", "all"):
        dump_raw(wb.create_sheet("Shifts"), shifts, ["Start", "End", "Created At"],
                 [26, 26, 26, 18, 10, 18, 9, 10, 18, 9, 9, 7])
    if RAW_TABS == "all":
        dump_raw(wb.create_sheet("Bookings"), book, ["Created At"],
                 [26, 18, 26, 13, 26, 26, 10])
        dump_raw(wb.create_sheet("Cancels"), cancel, ["Created At", "Start"],
                 [26, 18, 26, 18, 26, 18, 26, 10, 8, 9])

    # ------------------------------------------------------------- Cancel_Analysis (LIVE classification)
    ws = wb.create_sheet("Cancel_Analysis")
    ca = c_anchor[["Action ID", "Created At", "Shift ID", "Action", "Worker ID", "Start", "Facility ID", "Lead Time"]].reset_index(drop=True)
    ws.append(list(ca.columns) + ["Class", "IsOffense", "NoticeBucket"])
    header(ws, 1, 11)
    for row in ca.itertuples(index=False):
        ws.append(list(row))
    for r in range(2, NCA + 1):
        ws[f"B{r}"].number_format = DATEF
        ws[f"F{r}"].number_format = DATEF
        ws[f"I{r}"] = f'=IF(D{r}="NO_CALL_NO_SHOW","NCNS",IF(H{r}>=24,"early","late"))'
        ws[f"J{r}"] = f'=IF(I{r}="early",0,1)'
        ws[f"K{r}"] = f'=IF(D{r}<>"WORKER_CANCEL","",IF(H{r}<4,"<4h",IF(H{r}<24,"4-24h",IF(H{r}<72,"24-72h","72h+"))))'
        for col in "IJK":
            ws[f"{col}{r}"].font = ARIAL
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{NCA}"
    for j, w in enumerate([26, 18, 26, 18, 26, 18, 26, 10, 9, 9, 12], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.cell(row=NCA + 2, column=1,
            value="Anchor cancels only (events on the clean universe). Class / IsOffense / NoticeBucket are live formulas; the classification rule in cell form.").font = NOTE

    # ------------------------------------------------------------- Shift_Analysis (values)
    ws = wb.create_sheet("Shift_Analysis")
    dump_values(ws, shift_an, ["Start"], [], ["GrossRev"] if False else [],
                [26, 26, 26, 18, 10, 8, 9, 9, 8, 8, 11, 7, 11, 16, 15, 11],
                note=None)
    for r in range(2, NSA + 1):
        ws[f"M{r}"].number_format = MONEY
    notes = [
        "Derivation rules (one row per clean shift; Charge>0 AND Time>0):",
        "  Worked = 1 if Verified TRUE.  FacDeleted = 1 if Deleted flag = 1.  Empty = 1 if Worked=0 AND FacDeleted=0.",
        "  GrossRev = Charge x Time.  FinalCancelClass = class of this shift's LAST cancel event (early/late/NCNS), or 'none'.",
        "  FinalCancelLead = lead time of that final event.  StartMonth = YYYY-MM of shift start.",
        "  Reproduced by analysis/clipboard_reliability_analysis.py. Summary tabs reference these columns with live formulas.",
    ]
    for i, t in enumerate(notes):
        ws.cell(row=NSA + 2 + i, column=1, value=t).font = NOTE

    # ------------------------------------------------------------- Booking_Analysis (values)
    ws = wb.create_sheet("Booking_Analysis")
    dump_values(ws, book_an, ["Created At"], [], [],
                [26, 18, 26, 26, 10, 11, 13, 10, 10, 10, 14, 12, 12, 15, 15])
    notes = [
        "Derivation rules (one row per booking joinable to the clean universe; point-in-time, no leakage):",
        "  LastMinute = 1 if Lead Time < 24h.",
        "  PriorOffenses / PriorNCNS = count of this worker's late/NCNS (or NCNS) events with Created At strictly BEFORE this booking.",
        "  FailAfter = 1 if this worker late-cancels/NCNSes THIS shift at or after this booking time.  NCNSAfter = same, NCNS only.",
        "  AnyCancelAfter = 1 if any cancel by this worker on this shift at/after booking.  RescueClaim = 1 if this shift already had a cancel before this claim.",
        "  ShiftWorked / ShiftFinalClass = outcome and final class of the shift.  ShiftFailedNaive = 1 if final class is late/NCNS (the contaminated shift-level metric).",
        "  Reproduced by analysis/clipboard_reliability_analysis.py. Calc_Prediction and Calc_LastMin_Correction reference these columns with live formulas.",
    ]
    for i, t in enumerate(notes):
        ws.cell(row=NBA + 2 + i, column=1, value=t).font = NOTE

    # ---- Named ranges: give every column a plain-English name so the summary
    # formulas read like "=COUNTIFS(FinalClass,\"late\",ShiftWorked,1)" instead of
    # "=COUNTIFS(Shift_Analysis!$N$2:$N$35927, ...)". Legend is on the ReadMe tab.
    SA, BA, CAn = "Shift_Analysis", "Booking_Analysis", "Cancel_Analysis"
    named = {
        # Shift_Analysis (one row per clean shift)
        "FinalClass":      f"{SA}!$N$2:$N${NSA}",   # early / late / NCNS / none
        "ShiftWorked":     f"{SA}!$J$2:$J${NSA}",   # 1 if the shift was worked
        "FacDeleted":      f"{SA}!$K$2:$K${NSA}",   # 1 if the facility deleted it
        "ShiftEmpty":      f"{SA}!$L$2:$L${NSA}",   # 1 if it died empty
        "GrossRev":        f"{SA}!$M$2:$M${NSA}",   # charge x hours
        "FinalLead":       f"{SA}!$O$2:$O${NSA}",   # hours of notice at the final cancel
        "StartMonth":      f"{SA}!$P$2:$P${NSA}",   # YYYY-MM
        "Facility":        f"{SA}!$C$2:$C${NSA}",
        # Booking_Analysis (one row per booking on the clean universe)
        "AllBookings":     f"{BA}!$A$2:$A${NBA}",   # booking id, for counting
        "LastMinute":      f"{BA}!$F$2:$F${NBA}",   # 1 if claimed < 24h before start
        "PriorOffenses":   f"{BA}!$G$2:$G${NBA}",   # late/NCNS before this booking
        "PriorNCNS":       f"{BA}!$H$2:$H${NBA}",   # no-shows before this booking
        "DidFail":         f"{BA}!$I$2:$I${NBA}",   # 1 if this worker failed this booking
        "DidNoShow":       f"{BA}!$J$2:$J${NBA}",   # 1 if this worker no-showed it
        "DidCancelAfter":  f"{BA}!$K$2:$K${NBA}",   # 1 if any cancel after booking
        "IsRescueClaim":   f"{BA}!$L$2:$L${NBA}",   # 1 if shift already had a cancel
        "BkShiftWorked":   f"{BA}!$M$2:$M${NBA}",   # 1 if that shift ended up worked
        "ShiftFailedNaive":f"{BA}!$O$2:$O${NBA}",   # naive shift-level fail flag
        # Cancel_Analysis (anchor cancel events)
        "CxClass":         f"{CAn}!$I$2:$I${NCA}",  # early / late / NCNS
        "NoticeBucket":    f"{CAn}!$K$2:$K${NCA}",  # <4h / 4-24h / 24-72h / 72h+
        "CxWorker":        f"{CAn}!$E$2:$E${NCA}",
        "CxLead":          f"{CAn}!$H$2:$H${NCA}",  # hours of notice (negative = after start)
    }
    for nm, ref in named.items():
        wb.defined_names.add(DefinedName(nm, attr_text=ref))

    # summary formulas now reference the names
    S_CLS, S_WK, S_FD, S_EM, S_REV, S_LEAD, S_MON, S_FAC = (
        "FinalClass", "ShiftWorked", "FacDeleted", "ShiftEmpty", "GrossRev", "FinalLead", "StartMonth", "Facility")
    B_LM, B_PO, B_PN, B_FA, B_NA, B_AC, B_RC, B_SW, B_SFN = (
        "LastMinute", "PriorOffenses", "PriorNCNS", "DidFail", "DidNoShow",
        "DidCancelAfter", "IsRescueClaim", "BkShiftWorked", "ShiftFailedNaive")
    NB = "COUNTA(AllBookings)"
    CA_CLS, CA_NB, CA_WK, CA_LEAD = ("CxClass", "NoticeBucket", "CxWorker", "CxLead")

    # ------------------------------------------------------------- Calc_ShiftOutcomes
    ws = wb.create_sheet("Calc_ShiftOutcomes")
    r = block(ws, 1, "Refill outcomes by final cancel event (each shift counted once)",
              ["Final cancel type", "Shifts", "Refilled & worked", "Died empty", "Facility deleted", "Refill %", "Empty %"],
              [[lbl,
                f'=COUNTIF({S_CLS},"{k}")',
                f'=COUNTIFS({S_CLS},"{k}",{S_WK},1)',
                f'=COUNTIFS({S_CLS},"{k}",{S_EM},1)',
                f'=COUNTIFS({S_CLS},"{k}",{S_FD},1)',
                f"=C{i}/B{i}", f"=D{i}/B{i}"]
               for i, (lbl, k) in enumerate([("Early cancel (>=24h)", "early"),
                                             ("Late cancel (<24h)", "late"),
                                             ("No-call-no-show", "NCNS")], start=3)],
              pct=(6, 7), widths=[24, 10, 16, 12, 15, 10, 10])
    r = block(ws, r, "Where cancel-driven empty shifts come from",
              ["Source", "Empty shifts", "% of cancel-driven empties"],
              [["Late cancel", "=D4", f"=B{r+2}/B{r+5}"],
               ["NCNS", "=D5", f"=B{r+3}/B{r+5}"],
               ["Early cancel", "=D3", f"=B{r+4}/B{r+5}"],
               ["TOTAL", f"=SUM(B{r+2}:B{r+4})", 1]],
              pct=(3,), widths=[24, 12, 24])
    r = block(ws, r, "Late-cancel refill vs time remaining (runway)",
              ["Time left at final late cancel", "Shifts", "Refilled & worked", "Refill %", "Empty %"],
              [[lbl,
                f'=COUNTIFS({S_CLS},"late",{S_LEAD},">={lo}",{S_LEAD},"<{hi}")',
                f'=COUNTIFS({S_CLS},"late",{S_LEAD},">={lo}",{S_LEAD},"<{hi}",{S_WK},1)',
                f"=C{r+2+i}/B{r+2+i}",
                f'=COUNTIFS({S_CLS},"late",{S_LEAD},">={lo}",{S_LEAD},"<{hi}",{S_EM},1)/B{r+2+i}']
               for i, (lbl, lo, hi) in enumerate([("12-24h", 12, 24), ("4-12h", 4, 12), ("<4h", 0, 4)])],
              pct=(4, 5), widths=[26, 10, 16, 10, 10])
    months = ["2021-10", "2021-11", "2021-12", "2022-01"]
    r = block(ws, r, "Late-cancel refill by month (it was already improving; baseline honesty)",
              ["Month", "Late-cancelled shifts", "Refilled & worked", "Refill %"],
              [[m, f'=COUNTIFS({S_CLS},"late",{S_MON},"{m}")',
                f'=COUNTIFS({S_CLS},"late",{S_MON},"{m}",{S_WK},1)', f"=C{r+2+i}/B{r+2+i}"]
               for i, m in enumerate(months)],
              pct=(4,), widths=[10, 20, 16, 10])
    # No-show detection timing (the 'today' baseline for the reporting metric).
    # NCNS lead time is negative (logged after start); within 1h of start => lead >= -1.
    block(ws, r, "No-show detection timing today (why reactive recovery starts late)",
          ["Metric", "Events", "Share of no-shows"],
          [["No-show events (anchor)", f'=COUNTIF({CA_CLS},"NCNS")', 1],
           ["...logged within 1h of shift start (lead >= -1h)",
            f'=COUNTIFS({CA_CLS},"NCNS",{CA_LEAD},">=-1")', f"=B{r+3}/B{r+2}"],
           ["...logged within 4h of shift start (lead >= -4h)",
            f'=COUNTIFS({CA_CLS},"NCNS",{CA_LEAD},">=-4")', f"=B{r+4}/B{r+2}"]],
          pct=(3,), widths=[42, 10, 16],
          note="Lead time = shift start minus action time; negative means logged after start. This is the baseline the facility one-tap report is meant to lift toward 50%.")

    # ------------------------------------------------------------- Calc_Notice
    ws = wb.create_sheet("Calc_Notice")
    r = block(ws, 1, "Cancel events on the clean universe (anchor)",
              ["Type", "Events", "% of events"],
              [["Early (>=24h)", f'=COUNTIF({CA_CLS},"early")', "=B3/B$6"],
               ["Late (<24h)", f'=COUNTIF({CA_CLS},"late")', "=B4/B$6"],
               ["NCNS", f'=COUNTIF({CA_CLS},"NCNS")', "=B5/B$6"],
               ["TOTAL", "=SUM(B3:B5)", 1]],
              pct=(3,), widths=[16, 10, 12],
              note="Events exceed distinct cancelled shifts (a shift can be claimed→cancelled→re-claimed→cancelled); shift tables classify each shift once, by its final event.")
    block(ws, r, "Notice given (worker cancels on the clean universe)",
          ["Notice window", "Cancels", "% of worker cancels"],
          [[bkt, f'=COUNTIF({CA_NB},"{bkt}")', f"=B{r+2+i}/SUM(B{r+2}:B{r+5})"]
           for i, bkt in enumerate(["<4h", "4-24h", "24-72h", "72h+"])],
          pct=(3,), widths=[16, 10, 18])

    # ------------------------------------------------------------- Calc_Prediction
    ws = wb.create_sheet("Calc_Prediction")
    r = block(ws, 1, "Point-in-time prediction test (worker-level outcomes; engine room = Booking_Analysis)",
              ["Metric", "Value"],
              [["Bookings joinable to clean universe", f"={NB}"],
               ["Base fail rate (worker late-cancels/NCNSes after booking)", f"=SUM({B_FA})/B3"],
               ["Base NCNS rate", f"=SUM({B_NA})/B3"]],
              widths=[52, 14])
    ws["B4"].number_format = PCT
    ws["B5"].number_format = PCT
    tiers = []
    for i, (lbl, crit) in enumerate([("0", "0"), ("1", "1"), ("2", "2"), ("3+", '">=3"')]):
        rr = r + 2 + i
        tiers.append([lbl,
                      f"=COUNTIFS({B_PO},{crit})",
                      f"=COUNTIFS({B_PO},{crit},{B_FA},1)",
                      f"=C{rr}/B{rr}",
                      f"=D{rr}/$B$4",
                      f"=COUNTIFS({B_PO},{crit},{B_NA},1)/B{rr}"])
    r = block(ws, r, "Offense-history gradient (prior late/NCNS offenses, counted strictly before booking)",
              ["Prior offenses", "Bookings", "Failed after booking", "Fail %", "Lift vs base", "NCNS %"],
              tiers, pct=(4, 6), mult=(5,), widths=[14, 10, 18, 9, 12, 9])
    r = block(ws, r, "The 3+ prior-offense flag: coverage",
              ["Metric", "Value"],
              [["Share of bookings flagged", f'=COUNTIFS({B_PO},">=3")/B3'],
               ["Share of all failures caught", f'=COUNTIFS({B_PO},">=3",{B_FA},1)/SUM({B_FA})'],
               ["Share of NCNS caught", f'=COUNTIFS({B_PO},">=3",{B_NA},1)/SUM({B_NA})']],
              pct=(2,), widths=[52, 14])
    r = block(ws, r, "The first-timer wall (why prediction cannot be the strategy)",
              ["Metric", "Value"],
              [["Bookings ending in NCNS", f"=SUM({B_NA})"],
               ["...from workers with ZERO prior offenses", f"=COUNTIFS({B_NA},1,{B_PO},0)/B{r+2}"],
               ["...from workers with ZERO prior NCNS", f"=COUNTIFS({B_NA},1,{B_PN},0)/B{r+2}"]],
              widths=[52, 14])
    ws[f"B{r-2}"].number_format = PCT
    ws[f"B{r-1}"].number_format = PCT

    # ------------------------------------------------------------- Calc_LastMin_Correction
    ws = wb.create_sheet("Calc_LastMin_Correction")
    r = block(ws, 1, "Last-minute claims: naive (shift-level) vs corrected (worker-level) failure rates",
              ["Definition of 'failure'", "Base (all bookings)", "Last-minute claims", "Booked-ahead"],
              [["Shift-level: shift's final event is late/NCNS (naive; counts failures that predate the claim)",
                f"=SUM({B_SFN})/{NB}", f"=COUNTIFS({B_LM},1,{B_SFN},1)/SUM({B_LM})",
                f"=COUNTIFS({B_LM},0,{B_SFN},1)/({NB}-SUM({B_LM}))"],
               ["Worker-level: THIS claimant late-cancels/NCNSes after booking",
                f"=SUM({B_FA})/{NB}", f"=COUNTIFS({B_LM},1,{B_FA},1)/SUM({B_LM})",
                f"=COUNTIFS({B_LM},0,{B_FA},1)/({NB}-SUM({B_LM}))"],
               ["Worker-level: any cancellation after booking",
                f"=SUM({B_AC})/{NB}", f"=COUNTIFS({B_LM},1,{B_AC},1)/SUM({B_LM})",
                f"=COUNTIFS({B_LM},0,{B_AC},1)/({NB}-SUM({B_LM}))"]],
              pct=(2, 3, 4), widths=[58, 16, 16, 14],
              note="The naive definition brands last-minute claimers ~2x risky; the corrected one shows they are the MOST reliable segment. The gap is reverse causality; see rescue evidence below.")
    block(ws, r, "Rescue evidence, and the section-4 rescues-per-day arithmetic",
          ["Metric", "Value"],
          [["Last-minute claims (<24h before start)", f"=SUM({B_LM})"],
           ["...that were RESCUES (shift already had a cancel before the claim)", f"=COUNTIFS({B_LM},1,{B_RC},1)"],
           ["Rescue share of last-minute claims", f"=B{r+3}/B{r+2}"],
           ["Last-minute claims that held (no cancel of any kind after)", f"=1-COUNTIFS({B_LM},1,{B_AC},1)/B{r+2}"],
           ["Rescue claims where the shift ended up WORKED", f"=COUNTIFS({B_LM},1,{B_RC},1,{B_SW},1)/B{r+3}"],
           ["Failed shifts ALREADY recovered per day today (late+NCNS worked / 122)",
            f'=(COUNTIFS({S_CLS},"late",{S_WK},1)+COUNTIFS({S_CLS},"NCNS",{S_WK},1))/122'],
           ["Target ADDITIONAL saved shifts per day (increment; half-gap/month from Calc_Prize C13, / 30.4)", "=Calc_Prize!C13/30.4"],
           ["Target as a share of today's recoveries (the section-4 'half again')", f"=B{r+8}/B{r+7}"],
           ["Memo: last-minute rescue CLAIMS per day today (one channel only, 354/122)", f"=B{r+3}/122"]],
          widths=[64, 14],
          note="Section 4 compares like with like: today the marketplace already recovers about 7 failed shifts/day (row above the target), and the target adds ~3.4/day on top, i.e. roughly half again as many. The last-minute rescue CLAIMS figure (~2.9/day) is a different, narrower quantity and is shown only as a memo so the two are not confused.")
    for rr in (r + 4, r + 5, r + 6):
        ws[f"B{rr}"].number_format = PCT
    ws[f"B{r+7}"].number_format = "0.0"
    ws[f"B{r+8}"].number_format = "0.0"
    ws[f"B{r+9}"].number_format = PCT
    ws[f"B{r+10}"].number_format = "0.0"

    # ------------------------------------------------------------- Calc_Prize
    ws = wb.create_sheet("Calc_Prize")
    ws.cell(row=1, column=1, value="Prize sizing and 90-day target arithmetic").font = TITLE
    rows = [
        ("Take rate (case-provided input)", 0.22, "0%", True),
        ("Late-cancel empty shifts, 4-month window", f'=COUNTIFS({S_CLS},"late",{S_EM},1)', "#,##0", False),
        ("Gross facility charge destroyed (late empties)", f'=SUMIFS({S_REV},{S_CLS},"late",{S_EM},1)', MONEY, False),
        ("NCNS empty shifts, 4-month window", f'=COUNTIFS({S_CLS},"NCNS",{S_EM},1)', "#,##0", False),
        ("Gross facility charge destroyed (NCNS empties)", f'=SUMIFS({S_REV},{S_CLS},"NCNS",{S_EM},1)', MONEY, False),
        ("Combined CBH take destroyed (4 months, Cleveland)", "=(C4+C6)*C2", MONEY, False),
        ("Baseline month (Jan 2022): late-cancelled shifts", f'=COUNTIFS({S_CLS},"late",{S_MON},"2022-01")', "#,##0", False),
        ("Baseline month: late-cancel refill rate", f'=COUNTIFS({S_CLS},"late",{S_MON},"2022-01",{S_WK},1)/C8', "0.0%", False),
        ("Benchmark: early-cancel refill rate", f'=COUNTIFS({S_CLS},"early",{S_WK},1)/COUNTIF({S_CLS},"early")', "0.0%", False),
        ("Refill gap (benchmark - baseline)", "=C10-C9", "0.0%", False),
        ("Ceiling: full convergence (shifts/month) - NOT claimed", "=C8*C11", "#,##0", False),
        ("Target: close HALF the gap (shifts/month)", "=C12/2", "#,##0", False),
        ("Avg gross charge per late-cancelled shift", f'=AVERAGEIFS({S_REV},{S_CLS},"late")', MONEY, False),
        ("Monthly take protected at target (Cleveland)", "=C13*C14*C2", MONEY, False),
        ("Annualised take protected at target (Cleveland)", "=C15*12", MONEY, False),
        ("Avg gross charge per RESCUED shift (late/NCNS that ended worked)",
         f'=(SUMIFS({S_REV},{S_CLS},"late",{S_WK},1)+SUMIFS({S_REV},{S_CLS},"NCNS",{S_WK},1))'
         f'/(COUNTIFS({S_CLS},"late",{S_WK},1)+COUNTIFS({S_CLS},"NCNS",{S_WK},1))', MONEY, False),
        ("CBH take per rescued shift (this is the ~$68 cited in section 6)", "=C17*C2", MONEY, False),
    ]
    for i, (lbl, val, fmt, is_in) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=lbl).font = ARIAL
        c = ws.cell(row=i, column=3, value=val)
        c.font = BLUE if is_in else ARIAL
        if fmt:
            c.number_format = fmt
    ws.cell(row=len(rows) + 3, column=1,
            value="Every figure recalculates from Shift_Analysis. Column C. The half-gap target is the haircut number in the proposal; full convergence is a ceiling only.").font = NOTE
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["C"].width = 16

    # ------------------------------------------------------------- Sort_Facilities
    cls_map = clean.set_index("Shift ID")["FinalCancelClass"]
    emp_map = clean.set_index("Shift ID")["Empty"]
    fe = clean[(clean["FinalCancelClass"].isin(["late", "NCNS"])) & (clean["Empty"] == 1)]
    order = fe.groupby("Facility ID").size().sort_values(ascending=False).index.tolist()
    all_fac = order + sorted(set(clean["Facility ID"]) - set(order))
    ws = wb.create_sheet("Sort_Facilities")
    ws.cell(row=1, column=1, value="Facilities ranked by late/NCNS empty shifts (descending)").font = TITLE
    hd = ["Facility ID", "Late-cancel empties", "NCNS empties", "Total late/NCNS empties", "Share", "Cumulative share"]
    for j, h in enumerate(hd, start=1):
        ws.cell(row=2, column=j, value=h)
    header(ws, 2, len(hd))
    n = len(all_fac)
    for i, fac in enumerate(all_fac, start=3):
        ws.cell(row=i, column=1, value=fac).font = ARIAL
        ws.cell(row=i, column=2, value=f'=COUNTIFS({S_FAC},A{i},{S_CLS},"late",{S_EM},1)').font = ARIAL
        ws.cell(row=i, column=3, value=f'=COUNTIFS({S_FAC},A{i},{S_CLS},"NCNS",{S_EM},1)').font = ARIAL
        ws.cell(row=i, column=4, value=f"=B{i}+C{i}").font = ARIAL
        c = ws.cell(row=i, column=5, value=f"=D{i}/SUM($D$3:$D${n+2})"); c.font = ARIAL; c.number_format = PCT
        c = ws.cell(row=i, column=6, value=f"=SUM($D$3:D{i})/SUM($D$3:$D${n+2})"); c.font = ARIAL; c.number_format = PCT
    foot = n + 4
    ws.cell(row=foot, column=1, value="Facilities with >=1 late/NCNS empty:").font = BOLD
    ws.cell(row=foot, column=2, value=f'=COUNTIF(D3:D{n+2},">0")').font = ARIAL
    ws.cell(row=foot + 1, column=1, value="Top-10 facilities' share:").font = BOLD
    c = ws.cell(row=foot + 1, column=2, value=f"=SUM(D3:D12)/SUM(D3:D{n+2})"); c.font = ARIAL; c.number_format = PCT
    ws.freeze_panes = "A3"
    for j, w in enumerate([26, 17, 13, 20, 9, 15], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    # ------------------------------------------------------------- Sort_Workers_NCNS
    ncns_ev = c_anchor[c_anchor["Class"] == "NCNS"].groupby("Worker ID").size().sort_values(ascending=False)
    ws = wb.create_sheet("Sort_Workers_NCNS")
    ws.cell(row=1, column=1, value="Workers ranked by NCNS events on the clean universe (descending)").font = TITLE
    for j, h in enumerate(["Worker ID", "NCNS events", "Share", "Cumulative share"], start=1):
        ws.cell(row=2, column=j, value=h)
    header(ws, 2, 4)
    m = len(ncns_ev)
    for i, wid in enumerate(ncns_ev.index, start=3):
        ws.cell(row=i, column=1, value=wid).font = ARIAL
        ws.cell(row=i, column=2, value=f'=COUNTIFS({CA_WK},A{i},{CA_CLS},"NCNS")').font = ARIAL
        c = ws.cell(row=i, column=3, value=f"=B{i}/SUM($B$3:$B${m+2})"); c.font = ARIAL; c.number_format = PCT
        c = ws.cell(row=i, column=4, value=f"=SUM($B$3:B{i})/SUM($B$3:$B${m+2})"); c.font = ARIAL; c.number_format = PCT
    foot = m + 4
    k10 = max(1, int(-(-m // 10)))
    ws.cell(row=foot, column=1, value=f"Top 10% of these workers (n={k10}) cause this share of NCNS:").font = BOLD
    c = ws.cell(row=foot, column=2, value=f"=SUM(B3:B{2+k10})/SUM(B3:B{m+2})"); c.font = ARIAL; c.number_format = PCT
    ws.cell(row=foot + 1, column=1,
            value="Concentration is real, but Calc_Prediction shows 42% of NCNS bookings come from workers with zero prior offenses, so history can't carry a proactive strategy.").font = NOTE
    ws.freeze_panes = "A3"
    for j, w in enumerate([26, 12, 9, 15], start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    wb.save(OUT)
    print(f"written {OUT} | clean_shifts={len(shift_an)} anchor_cancels={len(c_anchor)} bookings={len(book_an)} "
          f"facilities={len(all_fac)} ncns_workers={len(ncns_ev)}")


if __name__ == "__main__":
    main()
