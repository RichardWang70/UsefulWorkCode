#!/usr/bin/env python3
"""
reprotect_sheets.py
───────────────────
Swaps the placeholder comment back to the original <sheetProtection …> tag
recorded in sheet_protection_log.jsonl.
"""

import json, shutil
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from tempfile import TemporaryDirectory
from collections import defaultdict

# ────────── USER OPTIONS ──────────
LOG_FILE  = Path("sheet_protection_log.jsonl")
OVERWRITE = True             # False → write *_reprotected.xlsx
PLACEHOLDER = "<!--SHEET_PROTECTION_REMOVED-->"
# ───────────────────────────────────

def load_log():
    mapping = defaultdict(list)   # {workbook: [(sheet_xml, tag), …]}
    if not LOG_FILE.exists():
        raise FileNotFoundError(LOG_FILE)
    with LOG_FILE.open(encoding="utf-8") as f:
        for line in f:
            if line.lstrip().startswith("#"):   # comment / header
                continue
            data = json.loads(line)
            mapping[Path(data["workbook"])].append((data["sheet_xml"],
                                                     data["tag"]))
    return mapping

def restore_tag(xml_path: Path, tag: str) -> bool:
    txt = xml_path.read_text(encoding="utf-8")
    if PLACEHOLDER not in txt:
        return False  # nothing to replace (maybe sheet deleted or already fixed)
    new_txt = txt.replace(PLACEHOLDER, tag, 1)
    xml_path.write_text(new_txt, encoding="utf-8")
    return True

def reprotect_wb(wb_path: Path, items) -> int:
    if not wb_path.exists():
        print(f"[!] missing: {wb_path}")
        return 0
    applied = 0
    with TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        with ZipFile(wb_path) as z:
            z.extractall(tmp)

        for sheet_xml, tag in items:
            xml_file = tmp / "xl" / "worksheets" / sheet_xml
            if not xml_file.exists():
                print(f"[!] {sheet_xml} missing in {wb_path.name}")
                continue
            if restore_tag(xml_file, tag):
                applied += 1

        if applied:
            out = wb_path if OVERWRITE else wb_path.with_stem(wb_path.stem + "_reprotected")
            with ZipFile(out, "w", ZIP_DEFLATED) as zout:
                for f in tmp.rglob("*"):
                    zout.write(f, f.relative_to(tmp))
            if OVERWRITE:
                shutil.copystat(wb_path, out)
    return applied

def main():
    work = load_log()
    total = 0
    for wb, items in work.items():
        done = reprotect_wb(wb, items)
        if done:
            print(f"[✓] {wb} – re-protected {done} sheet(s)")
            total += done
    print(f"Finished. {total} sheet(s) re-protected.")

if __name__ == "__main__":
    main()
