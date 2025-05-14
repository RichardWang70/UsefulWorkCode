#!/usr/bin/env python3
"""
unprotect_sheets.py
───────────────────
Replaces every <sheetProtection …> tag with a <!--SHEET_PROTECTION_REMOVED-->
placeholder and logs the original tag so it can be restored later.
"""

import re, json, shutil
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from tempfile import TemporaryDirectory
from datetime import datetime

# ────────── USER OPTIONS ──────────
ROOT_DIR  = Path(".")        # starting folder
RECURSIVE = True             # recurse into sub-folders?
OVERWRITE = True             # False → write *_unprotected.xlsx
LOG_FILE  = Path("sheet_protection_log.jsonl")
PLACEHOLDER = "<!--SHEET_PROTECTION_REMOVED-->"
# ───────────────────────────────────

RE_SHEET_PROT = re.compile(
    r"<sheetProtection\b[^>]*?(?:/>|></sheetProtection>)", re.I | re.S)

def iter_xlsx():
    pat = "**/*.xlsx" if RECURSIVE else "*.xlsx"
    return sorted(ROOT_DIR.glob(pat))

def log_entry(wb: Path, sheet_xml: str, tag: str):
    entry = {"workbook": str(wb.resolve()),
             "sheet_xml": sheet_xml,
             "tag": tag}
    with LOG_FILE.open("a", encoding="utf-8") as f:
        json.dump(entry, f); f.write("\n")

def strip_tag(xml_path: Path, wb: Path) -> int:
    txt = xml_path.read_text(encoding="utf-8")
    m = RE_SHEET_PROT.search(txt)
    if not m:
        return 0
    original_tag = m.group(0)
    new_txt = txt[:m.start()] + PLACEHOLDER + txt[m.end():]
    xml_path.write_text(new_txt, encoding="utf-8")
    log_entry(wb, xml_path.name, original_tag)
    return 1

def process_wb(path: Path) -> int:
    removed = 0
    with TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        with ZipFile(path) as z:
            z.extractall(tmp)
        for sxml in (tmp / "xl" / "worksheets").glob("sheet*.xml"):
            removed += strip_tag(sxml, path)

        if removed:
            out = path if OVERWRITE else path.with_stem(path.stem + "_unprotected")
            with ZipFile(out, "w", ZIP_DEFLATED) as zout:
                for f in tmp.rglob("*"):
                    zout.write(f, f.relative_to(tmp))
            if OVERWRITE:
                shutil.copystat(path, out)
    return removed

def main():
    LOG_FILE.write_text(
        f"# protection log – {datetime.now():%Y-%m-%d %H:%M:%S}\n",
        encoding="utf-8")
    total = 0
    for wb in iter_xlsx():
        c = process_wb(wb)
        if c:
            print(f"[✓] {wb} – removed {c} sheet(s)")
            total += c
    print(f"Done. {total} sheet(s) unprotected.")

if __name__ == "__main__":
    main()

