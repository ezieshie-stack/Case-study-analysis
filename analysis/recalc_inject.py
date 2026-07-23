"""
Recalculate an openpyxl-built workbook and inject cached values into its formula
cells — a pure-Python substitute for a LibreOffice recalc (headless LibreOffice
cannot load documents in this sandbox).

Approach: evaluate every formula with the `formulas` library, then edit each
worksheet's XML in place to add a <v> cached value to each <c> that has an <f>.
This preserves all styling written by openpyxl (styles live in separate parts),
unlike `formulas`' own writer.

Usage:  python analysis/recalc_inject.py <workbook.xlsx>
Exit 0 on success; non-zero if any formula evaluated to an Excel error.
"""

import logging
import shutil
import sys
import tempfile
import warnings
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

logging.disable(logging.WARNING)
warnings.filterwarnings("ignore")

import formulas  # noqa: E402

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ERR_TOKENS = {"#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#ERROR!"}


def scalar(v):
    """Extract a single Python scalar from a `formulas` Ranges/array value."""
    val = getattr(v, "value", v)
    try:
        import numpy as np
        if isinstance(val, np.ndarray):
            if val.size == 0:
                return None
            val = val.flat[0]
    except Exception:
        pass
    if isinstance(val, (list, tuple)):
        while isinstance(val, (list, tuple)):
            if not val:
                return None
            val = val[0]
    return val


def build_value_map(path):
    xl = formulas.ExcelModel().loads(str(path)).finish()
    sol = xl.calculate()
    out = {}
    errors = []
    for key, v in sol.items():
        # key like "'[file.xlsx]SHEETNAME'!A3"  (ranges contain ':' -> skip)
        if "!" not in key:
            continue
        ref = key.rsplit("!", 1)[1]
        if ":" in ref or "'" in ref:
            continue
        sheet = key.split("]", 1)[1].rsplit("'", 1)[0] if "]" in key else key.rsplit("!", 1)[0]
        sheet = sheet.strip("'").upper()
        s = scalar(v)
        out[(sheet, ref)] = s
        if isinstance(s, str) and s in ERR_TOKENS:
            errors.append(f"{sheet}!{ref} = {s}")
    return out, errors


def sheet_name_to_part(zpath):
    """Map upper-cased sheet display name -> worksheets/sheetN.xml part path."""
    with zipfile.ZipFile(zpath) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {}
    for rel in rels:
        rid_to_target[rel.get("Id")] = rel.get("Target")
    mapping = {}
    for sh in wb.find(f"{{{NS}}}sheets"):
        name = sh.get("name")
        rid = sh.get(f"{{{RELS_NS}}}id")
        target = rid_to_target.get(rid, "")
        if not target.startswith("/"):
            target = "xl/" + target.lstrip("/")
        else:
            target = target.lstrip("/")
        mapping[name.upper()] = target
    return mapping


def inject(path):
    path = Path(path)
    values, errors = build_value_map(path)

    part_map = sheet_name_to_part(path)
    ET.register_namespace("", NS)

    tmp = Path(tempfile.mkdtemp(prefix="inject_"))
    extract = tmp / "x"
    extract.mkdir()
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        z.extractall(extract)

    injected = 0
    for sheet_upper, part in part_map.items():
        xml_path = extract / part
        if not xml_path.exists():
            continue
        tree = ET.parse(xml_path)
        root = tree.getroot()
        for c in root.iter(f"{{{NS}}}c"):
            f = c.find(f"{{{NS}}}f")
            if f is None:
                continue
            ref = c.get("r")
            key = (sheet_upper, ref)
            if key not in values:
                continue
            val = values[key]
            if val is None:
                continue
            # drop any existing cached value
            old = c.find(f"{{{NS}}}v")
            if old is not None:
                c.remove(old)
            vel = ET.SubElement(c, f"{{{NS}}}v")
            if isinstance(val, bool):
                c.set("t", "b")
                vel.text = "1" if val else "0"
            elif isinstance(val, str):
                c.set("t", "str")
                vel.text = val
            else:
                if c.get("t") in ("str", "b", "e"):
                    del c.attrib["t"]
                num = float(val)
                vel.text = repr(int(num)) if num.is_integer() else repr(num)
            injected += 1
        tree.write(xml_path, xml_declaration=True, encoding="UTF-8")

    # force Excel to recompute on open as a belt-and-suspenders measure
    wb_xml_path = extract / "xl" / "workbook.xml"
    wb_tree = ET.parse(wb_xml_path)
    wb_root = wb_tree.getroot()
    calc = wb_root.find(f"{{{NS}}}calcPr")
    if calc is None:
        calc = ET.SubElement(wb_root, f"{{{NS}}}calcPr")
    calc.set("fullCalcOnLoad", "1")
    wb_tree.write(wb_xml_path, xml_declaration=True, encoding="UTF-8")

    out = path.with_suffix(".recalc.xlsx")
    if out.exists():
        out.unlink()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name in names:
            z.write(extract / name, name)
    shutil.move(str(out), str(path))
    shutil.rmtree(tmp, ignore_errors=True)

    print(f"injected cached values into {injected} formula cells")
    if errors:
        print(f"WARNING: {len(errors)} formula error(s):")
        for e in errors[:50]:
            print("  ", e)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(inject(sys.argv[1]))
