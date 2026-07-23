"""
Clipboard Health - Marketplace Reliability Case (Cleveland, Oct 2021 - Jan 2022)

Reproducible analysis: reads the three raw logs and writes every figure used in
the proposal to a single Excel workbook (deliverables/Clipboard_Analysis_Models.xlsx).

Inputs (not committed to this repo - place them in ./data/):
    data/Cleveland_shifts_logs.xlsx
    data/Booking_logs.xlsx
    data/Cancel_logs.xlsx

Definitions (also documented in the workbook's ReadMe tab):
    clean universe   shifts with Charge > 0 and Time > 0 (junk rows removed)
    worked           Verified == True (signed timesheet)
    facility-deleted Deleted == 1 (cancelled by the facility)
    empty            not worked AND not facility-deleted
    early cancel     WORKER_CANCEL with lead time >= 24h
    late cancel      WORKER_CANCEL with lead time < 24h
    NCNS             NO_CALL_NO_SHOW
    shift class      each shift classified ONCE, by its FINAL cancel event
    worker-level outcome  a booking "fails" only if THAT worker cancels THAT
                     shift late / no-shows AFTER the booking timestamp
    point-in-time    a worker's prior offenses are counted strictly before the
                     booking timestamp (no forward leakage)

Run:  python analysis/clipboard_reliability_analysis.py
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "deliverables" / "Clipboard_Analysis_Models.xlsx"

TAKE_RATE = 0.22
WEEKS_IN_WINDOW = 17.5  # Oct 1 2021 - Jan 31 2022 shift starts


def load():
    shifts = pd.read_excel(DATA / "Cleveland_shifts_logs.xlsx")
    book = pd.read_excel(DATA / "Booking_logs.xlsx").drop_duplicates(subset="Action ID")
    cancel = pd.read_excel(DATA / "Cancel_logs.xlsx").drop_duplicates(subset="Action ID")
    return shifts, book, cancel


def classify_cancel(row):
    if row["Action"] == "NO_CALL_NO_SHOW":
        return "NCNS"
    return "early" if row["Lead Time"] >= 24 else "late"


def main():
    shifts, book, cancel = load()

    # ---- cleaning ----------------------------------------------------------
    clean = shifts[(shifts["Charge"] > 0) & (shifts["Time"] > 0)].copy()
    clean["worked"] = clean["Verified"].fillna(False).astype(bool)
    clean["fac_deleted"] = clean["Deleted"] == 1.0
    clean["empty"] = ~clean["worked"] & ~clean["fac_deleted"]
    clean["rev"] = clean["Charge"] * clean["Time"]
    anchor_ids = set(clean["Shift ID"])
    m = clean.set_index("Shift ID")

    cancel["cls"] = cancel.apply(classify_cancel, axis=1)
    c_anchor = cancel[cancel["Shift ID"].isin(anchor_ids)].copy()

    tabs = {}

    tabs["1_Universe"] = pd.DataFrame(
        [
            ["Raw shifts", len(shifts)],
            ["Junk rows removed (Charge<=0 or Time<=0)", len(shifts) - len(clean)],
            ["Clean universe (anchor)", len(clean)],
            ["Distinct workers (raw)", shifts["Worker ID"].nunique()],
            ["Distinct facilities (clean)", clean["Facility ID"].nunique()],
            ["CNA+LVN share of clean shifts", round(clean["Agent Req"].isin(["CNA", "LVN"]).mean(), 3)],
            ["Worked (verified) share", round(clean["worked"].mean(), 3)],
            ["Facility-deleted share", round(clean["fac_deleted"].mean(), 3)],
            ["Empty (not worked, not deleted)", int(clean["empty"].sum())],
            ["Empty with no cancel event (mostly never claimed; out of scope)",
             int(clean.loc[~clean["Shift ID"].isin(set(c_anchor["Shift ID"])), "empty"].sum())],
            ["Empty following a worker cancel/NCNS (this case's scope)", None],  # filled below
            ["Shift-start window", f"{clean['Start'].min():%Y-%m-%d} to {clean['Start'].max():%Y-%m-%d}"],
            ["Avg gross facility charge per shift (Charge x Time)", round(clean["rev"].mean(), 2)],
        ],
        columns=["Metric", "Value"],
    )

    # ---- cancel events -----------------------------------------------------
    ev = c_anchor["cls"].value_counts()
    tabs["2_Cancel_Events"] = pd.DataFrame(
        {
            "Cancel type": ["Early (>=24h)", "Late (<24h)", "NCNS"],
            "Events": [ev.get("early", 0), ev.get("late", 0), ev.get("NCNS", 0)],
            "% of events": [round(ev.get(k, 0) / len(c_anchor) * 100, 1) for k in ["early", "late", "NCNS"]],
        }
    )
    wc = c_anchor[c_anchor["Action"] == "WORKER_CANCEL"]
    bins = pd.cut(wc["Lead Time"], [-np.inf, 4, 24, 72, np.inf], labels=["<4h", "4-24h", "24-72h", "72h+"])
    notice = (bins.value_counts(normalize=True).sort_index() * 100).round(1)
    tabs["2b_Notice_Given"] = pd.DataFrame(
        {"Notice window": notice.index.astype(str), "% of worker cancels": notice.values}
    )

    # ---- shift-level outcomes (final-event classification, exclusive) ------
    last_ev = c_anchor.sort_values("Created At").groupby("Shift ID").last()
    sh = last_ev[["cls", "Lead Time"]].join(m[["worked", "empty", "fac_deleted", "rev", "Start", "Facility ID"]], how="inner")
    tab3 = sh.groupby("cls").agg(
        shifts=("worked", "size"),
        refilled_worked=("worked", "sum"),
        died_empty=("empty", "sum"),
        facility_deleted=("fac_deleted", "sum"),
    ).reindex(["early", "late", "NCNS"])
    tab3["refill %"] = (tab3["refilled_worked"] / tab3["shifts"] * 100).round(1)
    tab3["empty %"] = (tab3["died_empty"] / tab3["shifts"] * 100).round(1)
    tabs["3_Shift_Outcomes"] = tab3.reset_index().rename(columns={"cls": "Final cancel type"})

    n_cancel_driven_empty = int(sh["empty"].sum())
    tabs["1_Universe"].loc[
        tabs["1_Universe"]["Metric"].str.startswith("Empty following"), "Value"
    ] = n_cancel_driven_empty

    e = sh[sh["empty"]]
    src = e["cls"].value_counts()
    tabs["4_Empty_Sources"] = pd.DataFrame(
        {
            "Source (final event)": ["Late cancel", "NCNS", "Early cancel", "TOTAL"],
            "Empty shifts": [src.get("late", 0), src.get("NCNS", 0), src.get("early", 0), int(src.sum())],
            "% of cancel-driven empties": [
                round(src.get("late", 0) / src.sum() * 100, 1),
                round(src.get("NCNS", 0) / src.sum() * 100, 1),
                round(src.get("early", 0) / src.sum() * 100, 1),
                100.0,
            ],
        }
    )

    # runway gradient for late cancels
    late_sh = sh[sh["cls"] == "late"]
    runway = pd.cut(late_sh["Lead Time"], [0, 4, 12, 24], labels=["<4h", "4-12h", "12-24h"])
    rt = late_sh.groupby(runway, observed=True).agg(n=("worked", "size"), refill=("worked", "mean"), empty=("empty", "mean"))
    tabs["5_Late_Runway"] = pd.DataFrame(
        {
            "Time left at late cancel": rt.index.astype(str),
            "Shifts": rt["n"].values,
            "Refill %": (rt["refill"] * 100).round(1).values,
            "Empty %": (rt["empty"] * 100).round(1).values,
        }
    )
    late_sh2 = late_sh.copy()
    late_sh2["month"] = late_sh2["Start"].dt.to_period("M").astype(str)
    mt = late_sh2.groupby("month").agg(n=("worked", "size"), refill=("worked", "mean"))
    tabs["5b_Late_Refill_Trend"] = pd.DataFrame(
        {"Month": mt.index, "Late-cancelled shifts": mt["n"].values, "Refill %": (mt["refill"] * 100).round(1).values}
    )

    # ---- booking behaviour: worker-level vs shift-level (the correction) ---
    b = book[book["Shift ID"].isin(anchor_ids)].copy()
    b["lastmin"] = b["Lead Time"] < 24

    c2 = cancel.merge(
        b[["Worker ID", "Shift ID", "Action ID", "Created At"]].rename(columns={"Action ID": "bid", "Created At": "bt"}),
        on=["Worker ID", "Shift ID"],
    )
    after = c2[c2["Created At"] >= c2["bt"]]
    b["cancel_after"] = b["Action ID"].isin(set(after["bid"]))
    b["fail_w"] = b["Action ID"].isin(set(after.loc[after["cls"].isin(["late", "NCNS"]), "bid"]))
    b["ncns_w"] = b["Action ID"].isin(set(after.loc[after["cls"] == "NCNS", "bid"]))

    fail_shifts_any = set(c_anchor.loc[c_anchor["cls"].isin(["late", "NCNS"]), "Shift ID"])
    b["fail_shift_level"] = b["Shift ID"].isin(fail_shifts_any)

    lm, ah = b[b["lastmin"]], b[~b["lastmin"]]
    tabs["6_LastMinute_Correction"] = pd.DataFrame(
        [
            ["Bookings joined to clean universe", len(b), "", ""],
            ["Last-minute claims (<24h before start)", len(lm), f"{len(lm)/len(b)*100:.1f}% of bookings", ""],
            ["SHIFT-level 'fail' rate (contaminated by pre-claim failures)",
             f"{b['fail_shift_level'].mean()*100:.1f}% base",
             f"last-minute {lm['fail_shift_level'].mean()*100:.1f}%",
             f"booked-ahead {ah['fail_shift_level'].mean()*100:.1f}%"],
            ["WORKER-level: this worker cancels late/NCNS AFTER booking",
             f"{b['fail_w'].mean()*100:.1f}% base",
             f"last-minute {lm['fail_w'].mean()*100:.1f}%",
             f"booked-ahead {ah['fail_w'].mean()*100:.1f}%"],
            ["WORKER-level: any cancel after booking",
             f"{b['cancel_after'].mean()*100:.1f}% base",
             f"last-minute {lm['cancel_after'].mean()*100:.1f}%",
             f"booked-ahead {ah['cancel_after'].mean()*100:.1f}%"],
            ["Conclusion", "Last-minute claimers are MORE reliable, not less.",
             "The shift-level 31% was reverse causality:", "rescue claims follow failures."],
        ],
        columns=["Metric", "Value", "Split A", "Split B"],
    )

    # rescue evidence
    first_cx = c_anchor.groupby("Shift ID")["Created At"].min()
    b2 = b.merge(first_cx.rename("first_cx"), left_on="Shift ID", right_index=True, how="left")
    b2["is_rescue"] = b2["first_cx"].notna() & (b2["Created At"] > b2["first_cx"])
    lm2 = b2[b2["lastmin"]]
    resc_ids = set(lm2.loc[lm2["is_rescue"], "Shift ID"])
    resc = m.loc[m.index.isin(resc_ids)]
    lmshift = m.loc[m.index.isin(set(lm2["Shift ID"]))]
    tabs["7_Rescue_Evidence"] = pd.DataFrame(
        [
            ["Last-minute claims (clean universe)", len(lm2)],
            ["...that were rescues (shift already had a cancel)", int(lm2["is_rescue"].sum())],
            ["Rescue share of last-minute claims", f"{lm2['is_rescue'].mean()*100:.1f}%"],
            ["Last-minute claims that held (no cancel after)", f"{(1-lm2['cancel_after'].mean())*100:.1f}%"],
            ["Rescued shifts that ended up worked", f"{resc['worked'].mean()*100:.1f}% (n={len(resc)})"],
            ["Any-last-minute-claim shifts that ended up worked", f"{lmshift['worked'].mean()*100:.1f}%"],
            ["Distinct workers claiming last-minute (clean universe)", lm2["Worker ID"].nunique()],
            ["Distinct workers claiming last-minute (full booking log)", book.loc[book["Lead Time"] < 24, "Worker ID"].nunique()],
        ],
        columns=["Metric", "Value"],
    )

    # ---- point-in-time prediction test (worker-level outcome) --------------
    off = cancel[cancel["cls"].isin(["late", "NCNS"])][["Worker ID", "Created At", "cls"]].sort_values("Created At")
    off_by_w = {w: g["Created At"].values for w, g in off.groupby("Worker ID")}
    ncns_by_w = {w: g["Created At"].values for w, g in off[off["cls"] == "NCNS"].groupby("Worker ID")}

    def prior(d, w, t):
        a = d.get(w)
        return 0 if a is None else int(np.searchsorted(a, t, side="left"))

    b["prior_off"] = [prior(off_by_w, w, t) for w, t in zip(b["Worker ID"], b["Created At"].values)]
    b["prior_ncns"] = [prior(ncns_by_w, w, t) for w, t in zip(b["Worker ID"], b["Created At"].values)]

    base = b["fail_w"].mean()
    tiers = pd.cut(b["prior_off"], [-1, 0, 1, 2, np.inf], labels=["0", "1", "2", "3+"])
    g = b.groupby(tiers, observed=True).agg(n=("fail_w", "size"), fail=("fail_w", "mean"), ncns=("ncns_w", "mean"))
    tabs["8_History_Gradient"] = pd.DataFrame(
        {
            "Prior late/NCNS offenses at claim (point-in-time)": g.index.astype(str),
            "Bookings": g["n"].values,
            "Fail % (late/NCNS after booking)": (g["fail"] * 100).round(1).values,
            "Lift vs 13.0% base": (g["fail"] / base).round(2).values,
            "NCNS %": (g["ncns"] * 100).round(1).values,
        }
    )
    t3 = b[b["prior_off"] >= 3]
    nb = b[b["ncns_w"]]
    tabs["9_Prediction_Test"] = pd.DataFrame(
        [
            ["Base fail rate (worker-level, late/NCNS after booking)", f"{base*100:.1f}%"],
            ["Base NCNS rate", f"{b['ncns_w'].mean()*100:.1f}%"],
            ["3+ prior offenses: fail rate", f"{t3['fail_w'].mean()*100:.1f}% = {t3['fail_w'].mean()/base:.2f}x lift (n={len(t3)})"],
            ["3+ flag: share of bookings flagged", f"{len(t3)/len(b)*100:.1f}%"],
            ["3+ flag: share of failures caught", f"{t3['fail_w'].sum()/b['fail_w'].sum()*100:.1f}%"],
            ["3+ flag: share of NCNS caught", f"{t3['ncns_w'].sum()/b['ncns_w'].sum()*100:.1f}%"],
            ["NCNS bookings with ZERO prior offenses (first-timers)", f"{(nb['prior_off']==0).mean()*100:.1f}% (n={len(nb)})"],
            ["NCNS bookings with ZERO prior NCNS", f"{(nb['prior_ncns']==0).mean()*100:.1f}%"],
            ["NCNS logged relative to shift start (median)", f"{c_anchor.loc[c_anchor['cls']=='NCNS','Lead Time'].median():.1f}h (i.e. ~35h AFTER start)"],
        ],
        columns=["Metric", "Value"],
    )

    # ---- prize + $ sizing --------------------------------------------------
    late_jan = late_sh2[late_sh2["month"] == "2022-01"]
    early_refill = tab3.loc["early", "refill %"]
    jan_refill = late_jan["worked"].mean() * 100
    late_empty_rev = sh.loc[(sh["cls"] == "late") & sh["empty"], "rev"]
    ncns_empty_rev = sh.loc[(sh["cls"] == "NCNS") & sh["empty"], "rev"]
    avg_rev_late = sh.loc[sh["cls"] == "late", "rev"].mean()
    jan_n = len(late_jan)
    ceiling_mo = jan_n * (early_refill - jan_refill) / 100
    half_mo = ceiling_mo / 2
    tabs["10_Prize_Sizing"] = pd.DataFrame(
        [
            ["Late-cancel empty shifts (window)", f"{int(src.get('late',0))}", f"gross ${late_empty_rev.sum():,.0f} | CBH take ${late_empty_rev.sum()*TAKE_RATE:,.0f}"],
            ["NCNS empty shifts (window)", f"{int(src.get('NCNS',0))}", f"gross ${ncns_empty_rev.sum():,.0f} | CBH take ${ncns_empty_rev.sum()*TAKE_RATE:,.0f}"],
            ["Combined 4-month cost, Cleveland only", f"{int(src.get('late',0)+src.get('NCNS',0))} shifts", f"~${(late_empty_rev.sum()+ncns_empty_rev.sum())*TAKE_RATE:,.0f} take"],
            ["Baseline (Jan 2022, trailing month)", f"{jan_n} late-cancelled shifts", f"refill {jan_refill:.1f}%"],
            ["Benchmark: early-cancel refill", f"{early_refill:.1f}%", "same marketplace, same shifts, more time"],
            ["Ceiling: late refill reaches early benchmark", f"~{ceiling_mo:.0f} shifts/month", "assumes full convergence - NOT claimed"],
            ["Target: close HALF the gap", f"~{half_mo:.0f} shifts/month (Cleveland)", f"~${half_mo*avg_rev_late*TAKE_RATE:,.0f}/mo take + churn protection"],
            ["Annualised (Cleveland only)", f"~{half_mo*12:,.0f} shifts", f"~${half_mo*12*avg_rev_late*TAKE_RATE:,.0f} take"],
        ],
        columns=["Item", "Shifts", "Value"],
    )

    # ---- facility exposure -------------------------------------------------
    fe = sh[sh["cls"].isin(["late", "NCNS"]) & sh["empty"]]
    concf = fe.groupby("Facility ID").size().sort_values(ascending=False)
    tabs["11_Facility_Exposure"] = pd.DataFrame(
        [
            ["Facilities with >=1 late/NCNS empty shift", f"{fe['Facility ID'].nunique()} of {clean['Facility ID'].nunique()}"],
            ["Top-10 facilities' share of late/NCNS empties", f"{concf.head(10).sum()/concf.sum()*100:.1f}%"],
            ["Late+NCNS shift-failures per week (window avg)", f"~{len(sh[sh['cls'].isin(['late','NCNS'])])/WEEKS_IN_WINDOW:.0f}"],
        ],
        columns=["Metric", "Value"],
    )

    readme = pd.DataFrame(
        {
            "Clipboard Health - Marketplace Reliability Case: analysis models": [
                "Market: Cleveland. Shift starts Oct 1 2021 - Jan 31 2022.",
                "Source: Cleveland_shifts_logs, Booking_logs, Cancel_logs (row-level).",
                "Generated by analysis/clipboard_reliability_analysis.py (fully reproducible).",
                "",
                "Definitions:",
                "  clean universe = shifts with Charge>0 and Time>0 (5,114 junk rows removed)",
                "  worked = Verified True (signed timesheet); facility-deleted = Deleted flag",
                "  empty = not worked AND not facility-deleted",
                "  early cancel >=24h notice; late <24h; NCNS = no-call-no-show",
                "  Each shift classified once, by its FINAL cancel event (exclusive classes).",
                "  Worker-level outcome: booking fails only if THAT worker late-cancels/NCNSes",
                "  THAT shift AFTER the booking timestamp (removes reverse-causality).",
                "  Prior offenses counted strictly before booking timestamp (point-in-time).",
                "",
                "Tab guide:",
                "  1  universe & cleaning        7  rescue-claim evidence",
                "  2  cancel events & notice     8  offense-history gradient",
                "  3  shift outcomes by type     9  prediction test summary",
                "  4  where empty shifts come from   10 prize & $ sizing",
                "  5  refill vs runway + trend   11 facility exposure",
                "  6  last-minute claim correction (shift- vs worker-level)",
            ]
        }
    )

    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        readme.to_excel(xw, sheet_name="0_ReadMe", index=False)
        for name, df in tabs.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
        # column widths
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 2, 80)

    print(f"written {OUT}")
    for name, df in tabs.items():
        print(f"\n### {name}\n{df.to_string(index=False)}")


if __name__ == "__main__":
    sys.exit(main())
