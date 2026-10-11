# Place Branding Writing Project (MK009A, Content Management I)

**Deadline:** 2 November 2026, via Blackboard (check the submission time there).
**Course contact:** Ahmed El Gody, ahmed.elgody@oru.se.
**Assignment brief:** `Place_Branding_Instructions.pdf` (the course PDF, kept here for reference).

This folder is the working kit for the assignment. It is organised in the order the work happens.

| File | What it is | Status |
|---|---|---|
| `01_Project_Plan.md` | Choosing the place and angle, scope decisions, week-by-week timeline to the deadline | Place chosen: the new restaurant, lecturer's confirmation pending |
| `02_Research_Guide.md` | Desk-research checklist, visit observation template, source log | Ready to use |
| `03_Interview_Guide.md` | Consent script, question bank, note template, quotation verification log | Ready to use |
| `07_Restaurant_Field_Guide.md` | Approach script, boxes to tick, and question sets for the owner, customers and neighbours at the chosen restaurant | Ready to use |
| `fieldwork/` | Where raw notes go: observation notes, interview notes, source log, photos list | Empty until fieldwork |
| `04_Communication_Plan.md` | The short communication plan, with course concepts mapped to each decision | Skeleton with placeholders |
| `texts/01_Feature.md` | Feature story based on the interview (counts as one of the four texts) | Structure and placeholders |
| `texts/02_Press_Release.md` | Press release for local journalists | Structure and placeholders |
| `texts/03_Webpage.md` | Practical webpage for residents who have not visited | Structure and placeholders |
| `texts/04_LinkedIn.md` | LinkedIn text for community organisations or partners | Structure and placeholders |
| `texts/05_Optional_Social_Caption.md` | Optional fifth text (social caption) | Optional |
| `05_References.md` | Reference list (APA 7th), Kavaratzis entry complete, Smith and Friedmann need the editions used | Partial |
| `06_Submission_Checklist.md` | Final checks before uploading | Ready to use |
| `build_portfolio.py` | Assembles plan, texts and references into one `.docx` and `.pdf` in `deliverables/` | Ready to use |

## How the templates work

Each template mixes two kinds of lines:

- **Submittable text**, written as normal Markdown.
- **Guidance**, written as blockquote lines beginning with `> ✎`. These explain what goes where and which course concept applies. `build_portfolio.py` strips every `> ✎` line, so they never reach the submitted document. Delete them by hand if you assemble the portfolio another way.

Placeholders look like `[[PLACE NAME]]`. The build script refuses to run while any `[[...]]` placeholder remains, so nothing unfinished slips into the submission.

## What must be true before submission

1. The place was actually visited (or the plan explains why not) and someone connected to it was actually interviewed, with permission.
2. Every quotation in the feature appears, word for word, in `fieldwork/interview_notes.md` with a timestamp or note reference.
3. Each text is labelled with **audience, format and intended platform**.
4. The communication plan names the course concepts it uses and cites the editions actually read.
5. The portfolio follows the Blackboard guidance on AI use. Use this kit as scaffolding and reporting discipline; the reporting, quotations and judgements must be your own.

## Build the portfolio

```bash
python3 place-branding/build_portfolio.py          # writes deliverables/Place_Branding_Portfolio.docx and .pdf
python3 place-branding/build_portfolio.py --check  # only reports remaining placeholders and guidance lines
```

Requires `pandoc` (installed in this environment). PDF output needs a LaTeX or `wkhtmltopdf` engine; if none is available the script writes the `.docx` only and says so.
