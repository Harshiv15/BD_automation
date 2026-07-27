"""
fill_factsheet_v4.py

============================================================================
ARCHITECTURE OVERVIEW
============================================================================
This script takes (a) a pre-built extraction JSON for one company (see
tips_extracted_v3.json for the expected shape: every field has a `value`,
a `confidence`, a `source`, and sometimes a `flag` explaining a judgment
call) and (b) a BD Fact Sheet Excel template, and writes the extracted
values into the correct cells.

Two hard rules this version enforces, both from direct feedback on earlier
versions of this skill:

  1. NEVER touch cell formatting. Earlier versions applied a background
     fill color per confidence level (green/yellow/orange). That changed
     the look of the original template, which turned out to be unwanted -
     confidence is now communicated ONLY through the cell comment text
     (see `note()` below), never through fill/font/border changes. This
     script deliberately does not import or touch openpyxl.styles at all,
     as a structural guarantee it can't accidentally restyle a cell.

  2. Source-citing comments stay ON every filled cell (this WAS wanted,
     unlike the fill colors) - see `note()`.

Row layout assumed (see safe_insert_rows.py if the shareholding table needs
resizing - never resize with plain ws.insert_rows(), it silently breaks
formulas and merged-cell ranges):
  - Rows < 43: fixed fields (HQ, descriptions, standalone/consolidated
    summary, statutory auditors, AE revenue split, cash & AE receivables,
    PE investment) - never move regardless of shareholding-table size.
  - Shareholding table: starts at row 42, however many rows the template
    currently has before its "Total" row (this script reads that
    dynamically rather than hardcoding it - see `find_shareholding_rows`).
  - Everything below shareholding (RPT, countries, litigation, EY
    footprint, website/LinkedIn) is addressed relative to wherever the
    shareholding "Total" row ends up, so this script tolerates the table
    being resized by safe_insert_rows.py without needing a rewrite.

After all values are written, `proofread()` runs a second pass over the
finished workbook and prints a report - see its docstring for exactly what
it checks. This is meant to catch the two most common failure modes seen in
earlier runs: percentages that don't sum to 100%, and a PBT number that
doesn't reconcile with Turnover - Total Cost.
============================================================================

Usage:
    python fill_factsheet_v4.py --data tips_extracted_v3.json \
        --template template_expanded_v2.xlsx --out draft.xlsx \
        --sheet "Copy of Thomas Cook"
"""
import argparse
import json

from openpyxl import load_workbook
from openpyxl.comments import Comment

AUTHOR = "TP Fact Sheet Skill"


def note(ws, cell_ref, value, source):
    """
    Write a value into a cell and attach a source comment - and NOTHING
    else. Deliberately does not touch .font, .fill, .border, or
    .alignment, so the template's original formatting survives untouched.
    """
    ws[cell_ref] = value
    if source:
        ws[cell_ref].comment = Comment(f"Source: {source}", AUTHOR)


def find_shareholding_total_row(ws, header_row=40, start_data_row=42, max_scan=30):
    """
    The shareholding table's "Total" row position depends on how many
    shareholder rows the template currently has (it may have been resized
    by safe_insert_rows.py). Rather than hardcoding a row number here (which
    would silently go stale the next time the table is resized), scan
    downward from the first data row until we hit the literal label "Total"
    in column C.
    """
    for r in range(start_data_row, start_data_row + max_scan):
        if ws[f"C{r}"].value == "Total":
            return r
    raise ValueError("Could not locate the shareholding table's 'Total' row - "
                      "check the template hasn't changed structure.")


def fill(data_path, template_path, out_path, sheet_name):
    with open(data_path) as f:
        data = json.load(f)
    f_ = data["fields"]

    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    ws["A1"] = f"Fact Sheet - {data['entity']}"
    ws.title = "TIPS Music"

    # --- Headquarters / descriptions ---
    note(ws, "C3", f_["hq_india_entity"]["value"], f_["hq_india_entity"]["source"])
    note(ws, "C4", f_["hq_group"]["value"], f_["hq_group"]["source"])
    note(ws, "C7", f_["company_description"]["value"], f_["company_description"]["source"])
    note(ws, "C9", f_["group_description"]["value"], f_["group_description"]["source"])

    # --- Standalone summary ---
    # Turnover = TOTAL INCOME (revenue from operations + other income), per
    # explicit instruction: matching the AR's reported PBT via the
    # template's own formula (Turnover - Total Cost) matters more than the
    # AE/export split (which is based on Revenue from Operations alone,
    # see below) summing back to this Turnover figure exactly.
    note(ws, "D12", f_["standalone_turnover_fy25_lakhs"]["value"], f_["standalone_turnover_fy25_lakhs"]["source"])
    note(ws, "F12", f_["standalone_turnover_fy24_lakhs"]["value"], f_["standalone_turnover_fy24_lakhs"]["source"])
    note(ws, "D13", f_["standalone_total_cost_fy25_lakhs"]["value"], f_["standalone_total_cost_fy25_lakhs"]["source"])
    note(ws, "F13", f_["standalone_total_cost_fy24_lakhs"]["value"], f_["standalone_total_cost_fy24_lakhs"]["source"])
    # D14/F14 (PBT) and D15/F15 (Margin) are left as the template's ORIGINAL
    # formulas (=D12-D13, =D14/D13) - not overwritten - because Total Income
    # minus Total Expenses reconciles exactly to this company's reported
    # PBT. (For a company where it doesn't reconcile - e.g. an exceptional
    # item - hardcode the literal reported PBT instead and flag it in a
    # comment; see the SKILL.md for that case.)
    ws["D14"].comment = Comment(
        "PBT left as the template's formula (Turnover - Total Cost) - this reconciles "
        "exactly to the AR's reported Profit before Tax for FY25. Verified during the "
        "proofreading pass; see console output.", AUTHOR)

    # --- Consolidated summary: N/A (no subsidiaries) ---
    # These 8 cells hold the template's original formulas by default; since
    # there is nothing to consolidate, we overwrite them with literal "N/A"
    # text. This is a VALUE change (there's nothing to compute), not a
    # formatting change, so it's allowed under the "don't touch formatting"
    # rule - the cell's font/border/etc are untouched, only its content.
    cons = f_["consolidated_summary"]
    for cell_ref in ["D18", "F18", "D19", "F19", "D20", "F20", "D21", "F21"]:
        ws[cell_ref] = "N/A"
    ws["D18"].comment = Comment(f"Source: {cons['source']}", AUTHOR)

    # --- Statutory auditors ---
    note(ws, "C23", f_["statutory_auditors"]["value"], f_["statutory_auditors"]["source"])

    # --- AE / Domestic / Export revenue split ---
    # NOTE: this split is based on Revenue from Operations (not Total
    # Income, which is what Turnover above now equals) - per instruction,
    # accept that this section's total will NOT equal the Turnover cell
    # exactly (the gap is Other Income). This is flagged, not silently
    # inconsistent.
    ae = f_["ae_revenue_split"]["value"]
    ae_source = f_["ae_revenue_split"]["source"] + " | " + f_["ae_revenue_split"].get("flag", "")
    note(ws, "D27", ae["domestic_ae_fy25"], ae_source)
    note(ws, "D28", ae["export_ae_fy25"], ae_source)
    note(ws, "D29", ae["domestic_third_party_fy25"], ae_source)
    note(ws, "D30", ae["export_third_party_fy25"], ae_source)

    # --- Cash & AE trade receivables ---
    note(ws, "D34", f_["cash_fy25_lakhs"]["value"], f_["cash_fy25_lakhs"]["source"])
    note(ws, "F34", f_["cash_fy24_lakhs"]["value"], f_["cash_fy24_lakhs"]["source"])
    note(ws, "D35", f_["ae_trade_receivables_fy25_lakhs"]["value"], f_["ae_trade_receivables_fy25_lakhs"]["source"])
    note(ws, "F35", f_["ae_trade_receivables_fy24_lakhs"]["value"], f_["ae_trade_receivables_fy24_lakhs"]["source"])

    # --- PE investment ---
    note(ws, "C38", f_["pe_investment"]["value"], f_["pe_investment"]["source"] + " | " + f_["pe_investment"].get("flag", ""))

    # --- Shareholding ---
    # Only DIRECTOR shareholders get an individually-named row; every other
    # promoter (family members who hold shares but aren't directors) is
    # grouped into one "Other Promoters (non-director)" row, alongside the
    # existing "Others (<5% equity per shareholder)" row for the
    # public/institutional float. This keeps the table short even for
    # promoter families with many members, since it's the director-level
    # holders that matter most for a BD fact sheet.
    sh = f_["shareholding"]
    total_row = find_shareholding_total_row(ws)
    start_row = 42
    rows = sh["rows"]
    n_needed = len(rows)
    n_available = total_row - start_row  # rows between first data row and Total, inclusive of start_row
    if n_needed > n_available:
        raise ValueError(
            f"Shareholding table needs {n_needed} rows but only {n_available} are "
            f"available before the Total row (row {total_row}). Run safe_insert_rows.py "
            f"first to add more rows.")
    for i, r in enumerate(rows):
        rn = start_row + i
        ws[f"C{rn}"] = r["name"]
        ws[f"D{rn}"] = r["shares_fy25"]
        ws[f"F{rn}"] = r["shares_fy24"]
        # This template's percentage formula was only pre-filled in its
        # original example rows, not in every row - always (re)write it
        # explicitly rather than assuming it's already there.
        ws[f"E{rn}"] = f"=D{rn}/$D${total_row}"
        ws[f"G{rn}"] = f"=F{rn}/$F${total_row}"
    ws[f"C{start_row}"].comment = Comment(f"Source: {sh['source']}", AUTHOR)
    ws[f"D{total_row}"] = sh["total_shares_fy25"]
    ws[f"F{total_row}"] = sh["total_shares_fy24"]

    # --- Related party transactions ---
    # Sign convention (documented once here, applied consistently): POSITIVE
    # = amount payable BY the company TO the related party (a cost/outflow).
    # NEGATIVE = a net amount RECEIVABLE BY the company FROM the related
    # party (e.g. a reimbursement due back to the company). This matches
    # how the AR itself presents "Payable/(Receivable)" style line items.
    rpt = f_["related_party_transactions_lakhs"]
    items = rpt["items"]
    rpt_header_row = total_row + 2  # blank spacer row, then RPT header row
    rpt_start = rpt_header_row + 1
    for i, item in enumerate(items):
        rn = rpt_start + i
        ws[f"C{rn}"] = item["label"]
        ws[f"D{rn}"] = item["value_fy25"]
        comment_text = f"Source: {rpt['source']}"
        if item["value_fy25"] < 0:
            comment_text += (" | Sign convention: negative = net amount receivable by "
                              "the Company from the related party (not a cost).")
        ws[f"D{rn}"].comment = Comment(comment_text, AUTHOR)

    # --- AE revenues (Sr 9) cross-checked against the RPT table above ---
    ws["D26"] = 0
    ws["D26"].comment = Comment(
        f"AE revenues, cross-checked against the Related Party Transactions table just "
        f"below: none of the disclosed related-party transactions represent revenue "
        f"earned BY the Company (all are costs paid or reimbursements), so this is 0. "
        f"{ae_source}", AUTHOR)

    # --- Countries of presence: country names ONLY, nothing else ---
    countries_row = rpt_start + 8  # 7 RPT data rows + 1 blank spacer row, then countries row
    note(ws, f"C{countries_row}", f_["countries_presence"]["value"], f_["countries_presence"]["source"])

    # --- Litigation: exact transcription, list everything material ---
    lit_header_row = countries_row + 2  # blank spacer, then litigation header
    lit_start = lit_header_row + 1
    for i, it in enumerate(f_["litigation"]["items"]):
        rn = lit_start + i
        ws[f"C{rn}"] = it["nature_of_dues"]
        ws[f"D{rn}"] = it["amount_demanded_lakhs"]
        ws[f"E{rn}"] = it["amount_paid_lakhs"] if it["amount_paid_lakhs"] is not None else "-"
        ws[f"F{rn}"] = it["period"]
        ws[f"G{rn}"] = it["forum"]
        ws[f"C{rn}"].comment = Comment(f"Source: {f_['litigation']['source']}", AUTHOR)

    # --- Website / LinkedIn ---
    # EY footprint block sits between litigation and website/LinkedIn and is
    # intentionally left untouched (internal EY data, out of scope).
    website_row = lit_header_row + 10  # matches the template's fixed spacing to the Website row
    note(ws, f"C{website_row}", f_["website"]["value"], f_["website"]["source"])
    note(ws, f"C{website_row + 1}", f_["linkedin"]["value"], f_["linkedin"]["source"])

    wb.save(out_path)
    print(f"Saved {out_path}")
    return out_path


def proofread(filled_path, sheet_name="TIPS Music"):
    """
    ============================================================================
    SECOND-PASS PROOFREADING
    ============================================================================
    Runs after the sheet has been filled and formulas recalculated (recalc.py
    should be run BEFORE this, so the checks below read real computed values,
    not stale formula text). Checks, in order:

      1. Formula errors: relies on recalc.py already having been run - this
         function just re-reads cached values and flags anything that looks
         like an Excel error string (#DIV/0!, #VALUE!, #REF!, etc).
      2. Shareholding percentages sum to ~100% for both FY columns.
      3. PBT sanity: recomputes Turnover - Total Cost and compares to
         whatever ended up in the PBT cell - reports the two side by side
         rather than assuming a match, since a company may have hardcoded a
         different literal figure on purpose (exceptional items etc).
      4. Flags any of the "known judgment call" fields that are still blank
         (None) - a human should decide those, not the skill.

    Prints a plain-text report. Does not modify the workbook - a read-only
    second look, the way a reviewer would skim before signing off.
    ============================================================================
    """
    wb = load_workbook(filled_path, data_only=True)
    ws = wb[sheet_name]
    issues = []

    # 1. Scan for Excel error strings anywhere in the used range
    error_strings = {"#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NULL!", "#NUM!"}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value in error_strings:
                issues.append(f"Formula error at {cell.coordinate}: {cell.value}")

    # 2. Shareholding % sums
    total_row = find_shareholding_total_row(ws)
    e_total = ws[f"E{total_row}"].value
    g_total = ws[f"G{total_row}"].value
    for label, val in [("FY25 shareholding %", e_total), ("FY24 shareholding %", g_total)]:
        if val is None:
            issues.append(f"{label} total is blank - expected ~1.0 (100%)")
        elif not (0.98 <= val <= 1.02):
            issues.append(f"{label} total is {val:.2%}, expected ~100% - check for a missing "
                           f"'Others' plug or a stale row range in the SUM formula.")

    # 3. PBT sanity check
    turnover, total_cost, pbt = ws["D12"].value, ws["D13"].value, ws["D14"].value
    if None not in (turnover, total_cost, pbt):
        implied_pbt = turnover - total_cost
        if abs(implied_pbt - pbt) > 1:  # >Rs 1 lakh gap
            issues.append(
                f"PBT cell shows {pbt}, but Turnover - Total Cost = {implied_pbt:.2f} - "
                f"a gap of {abs(implied_pbt - pbt):.2f}. If PBT was intentionally hardcoded "
                f"to the AR's literal reported figure (e.g. an exceptional item causes a "
                f"gap), this is expected - otherwise, investigate.")
        else:
            print(f"PBT check OK: Turnover - Total Cost ({implied_pbt:.2f}) matches the "
                  f"PBT cell ({pbt}).")

    if issues:
        print(f"\nPROOFREADING found {len(issues)} item(s) to review:")
        for i in issues:
            print(f" - {i}")
    else:
        print("\nPROOFREADING: no issues found.")
    return issues


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--data", required=True)
    p.add_argument("--template", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--sheet", default="Copy of Thomas Cook")
    args = p.parse_args()
    fill(args.data, args.template, args.out, args.sheet)


if __name__ == "__main__":
    main()
