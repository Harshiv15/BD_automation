"""
safe_insert_rows.py

============================================================================
ARCHITECTURE OVERVIEW
============================================================================
Problem this solves: openpyxl's built-in ws.insert_rows(idx, n) shifts cell
VALUES down by n rows, but does NOT rewrite formula text (so a formula like
"=D12-D13" sitting below the insertion point still says "=D12-D13" even
though the cells it should now point to have moved), and does NOT reliably
resize merged-cell ranges. It also only ever moves cells that already had a
*value* - a cell with formatting but no value is left behind, which is what
caused fonts/borders/alignment to look wrong in earlier runs of this skill
(a formatted-but-blank cell's style stayed at its old row while the row's
"neighbours" moved away from it).

How this script avoids that:
  1. Snapshot EVERY cell in the sheet's used range - value, formula-or-not,
     and every style attribute (font, fill, border, alignment, number
     format) - regardless of whether the cell currently holds a value. This
     is the key fix: styles on blank cells are preserved too.
  2. Snapshot row heights and column widths.
  3. Snapshot merged-cell ranges as plain strings (e.g. "B40:B45").
  4. Clear the sheet.
  5. Re-write every snapshotted cell at its NEW row (old_row + delta if
     old_row >= threshold, else unchanged), re-applying its exact style.
  6. For any formula string, shift every cell reference inside it the same
     way (regex over A1-style refs, respecting $ anchors) so formulas keep
     pointing at the *semantically* same cell, not the same literal address.
  7. Re-create merges and row heights at their new positions.

Net effect: the sheet looks and behaves exactly like the original, just
with `delta` extra blank (but correctly-styled) rows available starting at
`threshold`, ready for the fill script to write real values into.
============================================================================

Usage as a library:
    from safe_insert_rows import insert_rows_safe
    insert_rows_safe(src_path, out_path, sheet_name, threshold=43, delta=2)
"""
import re
from copy import copy

from openpyxl import load_workbook

# Matches A1-style cell references with optional $ column/row anchors,
# e.g. "D12", "$D$45", "F13". Used to find-and-shift references inside
# formula strings.
CELL_REF_RE = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')


def shift_formula(formula, threshold, delta):
    """Shift every row reference >= threshold by delta, inside a formula string."""
    def repl(m):
        col_abs, col, row_abs, row = m.groups()
        row_n = int(row)
        if row_n >= threshold:
            row_n += delta
        return f"{col_abs}{col}{row_abs}{row_n}"
    return CELL_REF_RE.sub(repl, formula)


def shift_ref(ref, threshold, delta):
    """Same shifting logic, applied to a plain range string like 'B40:B45'."""
    return shift_formula(ref, threshold, delta)


def insert_rows_safe(src_path, out_path, sheet_name, threshold, delta):
    wb = load_workbook(src_path, data_only=False)
    ws = wb[sheet_name]

    max_row = ws.max_row
    max_col = ws.max_column

    # --- 1. Snapshot merges, then remove them (can't move cells under a
    #     live merge without Excel complaining) ---
    old_merges = [str(r) for r in ws.merged_cells.ranges]
    for r in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(r))

    # --- 2. Snapshot row heights (indexed by row number) ---
    old_row_heights = {r: dim.height for r, dim in ws.row_dimensions.items() if dim.height is not None}

    # --- 3. Snapshot EVERY cell's value + full style, whether or not it
    #     currently holds a value. This is the fix for the earlier bug
    #     where blank-but-styled cells (e.g. border-only table cells) were
    #     never captured and so never moved with the rest of their row. ---
    snapshot = {}
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            snapshot[(cell.row, cell.column)] = {
                "value": cell.value,
                "font": copy(cell.font),
                "fill": copy(cell.fill),
                "border": copy(cell.border),
                "alignment": copy(cell.alignment),
                "number_format": cell.number_format,
                "protection": copy(cell.protection),
            }

    # --- 4. Clear all cell values (styles get overwritten in step 5 anyway,
    #     but clearing values first avoids any stale-formula evaluation
    #     issues while we rebuild) ---
    for row in ws.iter_rows(min_row=1, max_row=max_row + delta + 5, max_col=max_col):
        for cell in row:
            cell.value = None

    # --- 5. Re-write every cell at its shifted position with its original
    #     style, and shift any formula text so references stay correct ---
    for (r, c), info in snapshot.items():
        new_r = r + delta if r >= threshold else r
        cell = ws.cell(row=new_r, column=c)
        val = info["value"]
        if isinstance(val, str) and val.startswith("="):
            val = shift_formula(val, threshold, delta)
        cell.value = val
        cell.font = info["font"]
        cell.fill = info["fill"]
        cell.border = info["border"]
        cell.alignment = info["alignment"]
        cell.number_format = info["number_format"]
        cell.protection = info["protection"]

    # --- 6. Re-create row heights at shifted row numbers ---
    for r, height in old_row_heights.items():
        new_r = r + delta if r >= threshold else r
        ws.row_dimensions[new_r].height = height

    # --- 7. Re-create merges, shifted ---
    for m in old_merges:
        new_m = shift_ref(m, threshold, delta)
        ws.merge_cells(new_m)

    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--src", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--sheet", required=True)
    p.add_argument("--threshold", type=int, required=True)
    p.add_argument("--delta", type=int, required=True)
    args = p.parse_args()
    insert_rows_safe(args.src, args.out, args.sheet, args.threshold, args.delta)
    print(f"Saved {args.out}")
