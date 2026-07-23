# Case Study Analysis — Marketplace Reliability

Working repository for a marketplace-reliability case study (Cleveland shift data, Oct 2021 – Jan 2022): analysis of late cancellations and no-shows in a two-sided healthcare staffing marketplace, and a project proposal built on it.

## Contents

| Path | What it is |
|---|---|
| `deliverables/Clipboard_Case_Proposal.pdf` | The submission: 4-page narrative proposal + 2-page appendix |
| `deliverables/Clipboard_Analysis_Models.xlsx` | The Excel models: every figure in the proposal, one tab per analysis |
| `analysis/clipboard_reliability_analysis.py` | Reproducible script that generates the models workbook from the raw logs |
| `analysis/build_proposal.py` | Script that generates the proposal PDF |
| `analysis/Clipboard_Case_Analysis.md` | The reasoning record: findings, options considered, tests run, corrections made |

## Reproducing

The raw case data files are **not** committed (they are the case provider's materials). Place them in `data/` as `Cleveland_shifts_logs.xlsx`, `Booking_logs.xlsx`, `Cancel_logs.xlsx`, then:

```bash
pip install pandas openpyxl reportlab
python analysis/clipboard_reliability_analysis.py   # writes the models workbook
python analysis/build_proposal.py                   # writes the proposal PDF
```

## The one-line recommendation

When a booked shift fails (late cancellation or no-show), detect it within minutes and re-offer it instantly to proven same-day claimers — recovery, not punishment, is where the data says the leverage is.
