#!/usr/bin/env python3
"""
Extract a minimal JSON model from an .xlsx without external dependencies.

Outputs a JSON file that the static PWA can load:
- sheets: cell values + formulas
- inputs: labeled input cells to edit in the UI
- table: inferred part list table (rows/cols) to render

Usage:
  python3 scripts/extract_excel_model.py \\
    --xlsx /abs/path/to/TEMPLATE.xlsx \\
    --out  /abs/path/to/public/model.json
"""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from xml.etree import ElementTree as ET


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def num_to_col(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def split_addr(addr: str) -> Tuple[str, int]:
    m = CELL_RE.match(addr)
    if not m:
        raise ValueError(f"Bad cell ref: {addr}")
    return m.group(1), int(m.group(2))


def parse_shared_strings(z: zipfile.ZipFile) -> List[str]:
    try:
        xml = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(xml)
    out: List[str] = []
    for si in root.findall("main:si", NS):
        # plain <t> or rich text <r><t>
        t = si.find("main:t", NS)
        if t is not None and t.text is not None:
            out.append(t.text)
            continue
        parts: List[str] = []
        for r in si.findall("main:r", NS):
            rt = r.find("main:t", NS)
            if rt is not None and rt.text is not None:
                parts.append(rt.text)
        out.append("".join(parts))
    return out


def workbook_sheets(z: zipfile.ZipFile) -> List[Tuple[str, str]]:
    wb_root = ET.fromstring(z.read("xl/workbook.xml"))
    rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))

    rid_to_target: Dict[str, str] = {}
    for rel in rels_root.findall(
        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    ):
        rid_to_target[rel.attrib["Id"]] = rel.attrib["Target"]

    sheets: List[Tuple[str, str]] = []
    for sh in wb_root.findall("main:sheets/main:sheet", NS):
        name = sh.attrib.get("name") or ""
        rid = sh.attrib.get(f"{{{NS['r']}}}id")
        if not rid:
            continue
        target = rid_to_target.get(rid)
        if not target:
            continue
        sheets.append((name, "xl/" + target))
    return sheets


def parse_sheet_cells(
    z: zipfile.ZipFile, sheet_path: str, shared: List[str]
) -> Dict[str, Dict[str, Any]]:
    root = ET.fromstring(z.read(sheet_path))
    cells: Dict[str, Dict[str, Any]] = {}
    for c in root.findall(".//main:sheetData/main:row/main:c", NS):
        addr = c.attrib.get("r")
        if not addr:
            continue

        t = c.attrib.get("t")
        f_el = c.find("main:f", NS)
        v_el = c.find("main:v", NS)

        formula = (f_el.text or "").strip() if f_el is not None else None

        val: Optional[Any] = None
        if v_el is not None and v_el.text is not None:
            raw = v_el.text
            if t == "s":
                try:
                    val = shared[int(raw)]
                except Exception:
                    val = raw
            else:
                # keep as string; consumer may parse to number
                val = raw

        if t == "inlineStr":
            it = c.find("main:is/main:t", NS)
            if it is not None and it.text is not None:
                val = it.text

        # Normalize blanks
        if val == "":
            val = None

        cells[addr] = {"v": val, "f": formula}
    return cells


def guess_header_row(cells: Dict[str, Dict[str, Any]], max_scan_rows: int = 30) -> int:
    # heuristic: within first N rows, row with max textual values
    score_by_row: Dict[int, int] = {}
    for addr, cell in cells.items():
        col, row = split_addr(addr)
        if row <= 0 or row > max_scan_rows:
            continue
        v = cell.get("v")
        if v is None:
            continue
        score_by_row[row] = score_by_row.get(row, 0) + (2 if isinstance(v, str) else 1)
    if not score_by_row:
        return 1
    return max(score_by_row.items(), key=lambda kv: kv[1])[0]


def infer_table_start(sheet_name: str, cells: Dict[str, Dict[str, Any]]) -> int:
    # Use observed pattern: KITCHEN starts at row 10, WARDROBE at row 9.
    # Fallback: find first row where A{row} is a known part label.
    known = {"TOP", "BOTTOM", "RIGHT", "LEFT", "SHUTTER", "BACK", "SKERTING", "DUMMY"}
    candidates = []
    for addr, cell in cells.items():
        if not addr.startswith("A"):
            continue
        _, row = split_addr(addr)
        v = cell.get("v")
        if isinstance(v, str) and v.strip().upper() in known:
            candidates.append(row)
    if candidates:
        return min(candidates)
    # last resort: header heuristic + 1
    return guess_header_row(cells) + 1


def infer_table_end(
    cells: Dict[str, Dict[str, Any]], start_row: int, max_gap: int = 5
) -> int:
    # scan down col A until a few consecutive blanks
    gap = 0
    r = start_row
    last_nonblank = start_row
    # safe bound
    max_row = max(split_addr(a)[1] for a in cells.keys()) if cells else start_row
    while r <= max_row:
        a = cells.get(f"A{r}", {}).get("v")
        if isinstance(a, str) and a.strip() != "":
            last_nonblank = r
            gap = 0
        else:
            gap += 1
            if gap >= max_gap:
                break
        r += 1
    return last_nonblank


def infer_table_columns(
    cells: Dict[str, Dict[str, Any]], start_row: int, end_row: int
) -> List[str]:
    # choose columns that have any content/formula within the table region
    used_cols: set[str] = set()
    for addr, cell in cells.items():
        col, row = split_addr(addr)
        if row < start_row or row > end_row:
            continue
        if cell.get("v") is None and not cell.get("f"):
            continue
        used_cols.add(col)

    # keep a sane, Excel-like order. Prefer A..P and include O if present.
    # Also avoid super-wide sheets; cap to A..P (16 cols) for UI.
    max_col = col_to_num("P")
    ordered = [num_to_col(i) for i in range(1, max_col + 1) if num_to_col(i) in used_cols]
    return ordered


def labeled_inputs_for_sheet(sheet_name: str) -> List[Dict[str, Any]]:
    # Based on observed template structure (top area labels/values)
    base = [
        {"label_cell": "A1", "cell": "B1", "type": "number"},
        {"label_cell": "A2", "cell": "B2", "type": "number"},
        {"label_cell": "A3", "cell": "B3", "type": "number"},
        {"label_cell": "A4", "cell": "D4", "type": "number"},
        {"label_cell": "A5", "cell": "D5", "type": "number"},
        {"label_cell": "A6", "cell": "B6", "type": "number"},
        {"label_cell": "A7", "cell": "B7", "type": "number"},
        {"label_cell": "A8", "cell": "B8", "type": "number"},
        {"label_cell": "C1", "cell": "D1", "type": "text"},
        {"label_cell": "C2", "cell": "D2", "type": "text"},
    ]
    if sheet_name.strip().upper() == "WARDROBE":
        base += [
            {"label_cell": "A22", "cell": "B22", "type": "number"},
            {"label_cell": "C22", "cell": "D22", "type": "text"},
            {"label_cell": "A24", "cell": "B24", "type": "number"},
        ]
    return base


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True, help="Path to .xlsx file")
    ap.add_argument("--out", required=True, help="Output JSON path")
    args = ap.parse_args()

    xlsx = Path(args.xlsx)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(xlsx, "r") as z:
        shared = parse_shared_strings(z)
        sheets = workbook_sheets(z)
        model: Dict[str, Any] = {"source": str(xlsx.name), "sheets": {}}

        for sheet_name, sheet_path in sheets:
            cells = parse_sheet_cells(z, sheet_path, shared)
            start_row = infer_table_start(sheet_name, cells)
            end_row = infer_table_end(cells, start_row)
            cols = infer_table_columns(cells, start_row, end_row)

            # attach labels for inputs
            inputs = []
            for inp in labeled_inputs_for_sheet(sheet_name):
                label = cells.get(inp["label_cell"], {}).get("v")
                inputs.append(
                    {
                        "label": str(label).strip() if label is not None else inp["cell"],
                        "cell": inp["cell"],
                        "type": inp["type"],
                    }
                )

            model["sheets"][sheet_name] = {
                "cells": cells,
                "inputs": inputs,
                "table": {"startRow": start_row, "endRow": end_row, "columns": cols},
            }

    out.write_text(json.dumps(model, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {out} ({out.stat().st_size} bytes)")


if __name__ == "__main__":
    main()


