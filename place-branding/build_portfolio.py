#!/usr/bin/env python3
"""Assemble the place-branding portfolio into one document.

Reads the communication plan, the texts and the reference list, strips the
`> ✎` guidance lines, refuses to build while `[[...]]` placeholders remain, and
writes deliverables/Place_Branding_Portfolio.docx (and .pdf when a PDF engine
is available) using pandoc.

Usage:
    python3 place-branding/build_portfolio.py            # build
    python3 place-branding/build_portfolio.py --check    # report only
    python3 place-branding/build_portfolio.py --force    # build despite placeholders (draft review)
"""
import re
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
OUT_DIR = ROOT / "deliverables"
OUT_BASE = "Place_Branding_Portfolio"

ORDER = [
    HERE / "04_Communication_Plan.md",
    HERE / "texts" / "01_Feature.md",
    HERE / "texts" / "02_Press_Release.md",
    HERE / "texts" / "03_Webpage.md",
    HERE / "texts" / "04_LinkedIn.md",
    HERE / "texts" / "05_Optional_Social_Caption.md",  # dropped automatically if still all placeholders
    HERE / "05_References.md",
]

GUIDANCE = re.compile(r"^\s*>\s*✎")
PLACEHOLDER = re.compile(r"\[\[[^\]]*\]\]")


def clean(text: str) -> str:
    lines = [ln for ln in text.splitlines() if not GUIDANCE.match(ln)]
    # Collapse runs of blank lines left behind by stripped guidance.
    out, blank = [], 0
    for ln in lines:
        if ln.strip():
            blank = 0
            out.append(ln)
        else:
            blank += 1
            if blank <= 1:
                out.append(ln)
    return "\n".join(out).strip() + "\n"


def is_untouched_optional(path: Path, text: str) -> bool:
    """The optional fifth text is skipped if nothing in it has been filled in."""
    if path.name != "05_Optional_Social_Caption.md":
        return False
    body = clean(text)
    body_wo_placeholders = PLACEHOLDER.sub("", body)
    # If removing placeholders leaves only headings, labels and hashtags, treat as untouched.
    substantive = [ln for ln in body_wo_placeholders.splitlines()
                   if ln.strip() and not ln.startswith("#") and not ln.startswith("**") and ln.strip() != "---"]
    return len(substantive) == 0


def main(argv):
    check_only = "--check" in argv
    force = "--force" in argv

    sections, problems = [], []
    for path in ORDER:
        if not path.exists():
            problems.append(f"missing file: {path.relative_to(ROOT)}")
            continue
        raw = path.read_text(encoding="utf-8")
        if is_untouched_optional(path, raw):
            print(f"skip   {path.relative_to(ROOT)} (optional text, not filled in)")
            continue
        body = clean(raw)
        for m in PLACEHOLDER.finditer(body):
            line_no = body.count("\n", 0, m.start()) + 1
            problems.append(f"{path.relative_to(ROOT)}:{line_no}: placeholder {m.group(0)[:60]}")
        guidance_left = [i + 1 for i, ln in enumerate(raw.splitlines()) if "✎" in ln and not GUIDANCE.match(ln)]
        for ln in guidance_left:
            problems.append(f"{path.relative_to(ROOT)}:{ln}: guidance marker outside a '> ✎' line (would be published)")
        sections.append(body)
        print(f"read   {path.relative_to(ROOT)} ({len(body.split())} words after stripping guidance)")

    if problems:
        print(f"\n{len(problems)} item(s) still to resolve:")
        for p in problems:
            print("  " + p)
    else:
        print("\nNo placeholders or stray guidance left.")

    if check_only:
        return 1 if problems else 0
    if problems and not force:
        print("\nNot building. Fix the items above, or pass --force for a draft build.")
        return 1

    if shutil.which("pandoc") is None:
        print("pandoc not found; install it or assemble the files by hand.")
        return 2

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    combined = HERE / ".portfolio_combined.md"
    combined.write_text("\n\n\\newpage\n\n".join(sections), encoding="utf-8")

    docx = OUT_DIR / f"{OUT_BASE}.docx"
    subprocess.run(["pandoc", str(combined), "-o", str(docx), "--from", "markdown"], check=True)
    print(f"wrote  {docx.relative_to(ROOT)}")

    pdf = OUT_DIR / f"{OUT_BASE}.pdf"
    engines = [e for e in ("xelatex", "pdflatex", "wkhtmltopdf", "weasyprint", "typst") if shutil.which(e)]
    if engines:
        try:
            subprocess.run(["pandoc", str(combined), "-o", str(pdf), f"--pdf-engine={engines[0]}"], check=True)
            print(f"wrote  {pdf.relative_to(ROOT)} (engine: {engines[0]})")
        except subprocess.CalledProcessError as e:
            print(f"PDF build failed with {engines[0]}: {e}. The .docx is complete; export PDF from Word/LibreOffice.")
    else:
        print("No PDF engine found; the .docx is complete. Export PDF from Word or LibreOffice if Blackboard needs one.")
    combined.unlink(missing_ok=True)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
