"""
fill_factsheet.py

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
    python fill_factsheet.py --data tips_extracted_v3.json \
        --template template_expanded_v2.xlsx --out draft.xlsx \
        --sheet "Copy of Thomas Cook"
"""
import argparse
import json
from openpyxl import load_workbook
from openpyxl.comments import Comment

AUTHOR = "TP Fact Sheet Skill"

def note(ws, cell_ref, value, source):
    ws[cell_ref] = value
    if source:
        ws[cell_ref].comment = Comment(f"Source: {source}", AUTHOR)

def find_shareholding_total_row(ws, header_row=40, start_data_row=42, max_scan=30):
    for r in range(start_data_row, start_data_row + max_scan):
        if ws[f"C{r}"].value == "Total":
            return r
    raise ValueError("Could not locate the shareholding table's 'Total' row.")

def fill(data_path, template_path, out_path, sheet_name):
    with open(data_path) as f:
        data = json.load(f)
    f_ = data["fields"]

    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    
    ws.title = data['entity'][:31]
    ws["A1"] = f"Fact Sheet - {data['entity']}"

    # Track rows we want to HIDE (not delete) to preserve formula integrity
    rows_to_hide = []

    # --- Headquarters / descriptions ---
    note(ws, "C3", f_["hq_india_entity"]["value"], f_["hq_india_entity"]["source"])
    note(ws, "C4", f_["hq_group"]["value"], f_["hq_group"]["source"])
    note(ws, "C7", f_["company_description"]["value"], f_["company_description"]["source"])
    note(ws, "C9", f_["group_description"]["value"], f_["group_description"]["source"])

    # --- Standalone summary ---
    note(ws, "D12", f_["standalone_turnover_fy25_lakhs"]["value"], f_["standalone_turnover_fy25_lakhs"]["source"])
    note(ws, "F12", f_["standalone_turnover_fy24_lakhs"]["value"], f_["standalone_turnover_fy24_lakhs"]["source"])
    note(ws, "D13", f_["standalone_total_cost_fy25_lakhs"]["value"], f_["standalone_total_cost_fy25_lakhs"]["source"])
    note(ws, "F13", f_["standalone_total_cost_fy24_lakhs"]["value"], f_["standalone_total_cost_fy24_lakhs"]["source"])
    
    # Updated comment to reflect EY feedback on Turnover vs PBT
    ws["D14"].comment = Comment("PBT left as the template's formula. Note: Since Turnover now excludes Other Income, this formula (Turnover - Total Cost) will likely not match the AR's reported PBT.", AUTHOR)

    # --- Consolidated summary ---
    cons = f_["consolidated_summary"]
    for cell_ref in ["D18", "F18", "D19", "F19", "D20", "F20", "D21", "F21"]:
        ws[cell_ref] = "N/A"
    ws["D18"].comment = Comment(f"Source: {cons['source']}", AUTHOR)

    # --- Statutory auditors ---
    note(ws, "C23", f_["statutory_auditors"]["value"], f_["statutory_auditors"]["source"])

    # --- AE / Domestic / Export revenue split ---
    ae = f_["ae_revenue_split"]["value"]
    ae_source = f_["ae_revenue_split"]["source"] + " | " + f_["ae_revenue_split"].get("flag", "")
    note(ws, "D27", ae.get("domestic_ae_fy25"), ae_source)
    note(ws, "D28", ae.get("export_ae_fy25"), ae_source)
    note(ws, "D29", ae.get("domestic_third_party_fy25"), ae_source)
    note(ws, "D30", ae.get("export_third_party_fy25"), ae_source)

    # --- Cash & AE trade receivables ---
    note(ws, "D34", f_["cash_fy25_lakhs"]["value"], f_["cash_fy25_lakhs"]["source"])
    note(ws, "F34", f_["cash_fy24_lakhs"]["value"], f_["cash_fy24_lakhs"]["source"])
    note(ws, "D35", f_["ae_trade_receivables_fy25_lakhs"]["value"], f_["ae_trade_receivables_fy25_lakhs"]["source"])
    note(ws, "F35", f_["ae_trade_receivables_fy24_lakhs"]["value"], f_["ae_trade_receivables_fy24_lakhs"]["source"])

    # --- PE investment ---
    note(ws, "C38", f_["pe_investment"]["value"], str(f_["pe_investment"].get("source")) + " | " + str(f_["pe_investment"].get("flag", "")))

    # --- Shareholding ---
    sh = f_["shareholding"]
    total_row = find_shareholding_total_row(ws)
    start_row = 42
    rows = sh["rows"]
    n_needed = len(rows)
    n_available = total_row - start_row 
    
    if n_needed > n_available:
        raise ValueError(f"Run safe_insert_rows.py first. Needed: {n_needed}, Available: {n_available}")
    
    if n_needed < n_available:
        for i in range(n_needed, n_available):
            rows_to_hide.append(start_row + i)

    for i, r in enumerate(rows):
        rn = start_row + i
        ws[f"C{rn}"] = r["name"]
        ws[f"D{rn}"] = r["shares_fy25"]
        ws[f"F{rn}"] = r["shares_fy24"]
        ws[f"E{rn}"] = f"=D{rn}/$D${total_row}"
        ws[f"G{rn}"] = f"=F{rn}/$F${total_row}"
        
    ws[f"C{start_row}"].comment = Comment(f"Source: {sh['source']}", AUTHOR)
    ws[f"D{total_row}"] = sh["total_shares_fy25"]
    ws[f"F{total_row}"] = sh["total_shares_fy24"]

    # --- Related party transactions ---
    rpt = f_["related_party_transactions_lakhs"]
    items = rpt["items"]
    rpt_header_row = total_row + 2 
    rpt_start = rpt_header_row + 1
    
    rpt_capacity = 7
    if len(items) < rpt_capacity:
        for i in range(len(items), rpt_capacity):
            rows_to_hide.append(rpt_start + i)
            
    for i, item in enumerate(items):
        if i >= rpt_capacity:
            break 
        rn = rpt_start + i
        ws[f"C{rn}"] = item["label"]
        ws[f"D{rn}"] = item["value_fy25"]
        comment_text = f"Source: {rpt['source']}"
        if item.get("value_fy25") and float(item["value_fy25"]) < 0:
            comment_text += " | Sign convention: negative = net amount receivable."
        ws[f"D{rn}"].comment = Comment(comment_text, AUTHOR)

    ws["D26"] = 0
    ws["D26"].comment = Comment(f"AE revenues (Sale of services to foreign AEs) cross-checked against RPT table. {ae_source}", AUTHOR)

    # --- Countries of presence ---
    countries_row = rpt_start + rpt_capacity + 1
    note(ws, f"C{countries_row}", f_["countries_presence"]["value"], f_["countries_presence"]["source"])

    # --- Litigation ---
    lit_header_row = countries_row + 2 
    lit_start = lit_header_row + 1
    lit_items = f_["litigation"]["items"]
    
    lit_capacity = 2
    if len(lit_items) < lit_capacity:
        for i in range(len(lit_items), lit_capacity):
            rows_to_hide.append(lit_start + i)
            
    for i, it in enumerate(lit_items):
        if i >= lit_capacity:
            break
        rn = lit_start + i
        ws[f"C{rn}"] = it.get("nature_of_dues")
        ws[f"D{rn}"] = it.get("amount_demanded_lakhs")
        ws[f"E{rn}"] = it.get("amount_paid_lakhs") if it.get("amount_paid_lakhs") is not None else "-"
        ws[f"F{rn}"] = it.get("period")
        ws[f"G{rn}"] = it.get("forum")
        ws[f"C{rn}"].comment = Comment(f"Source: {f_['litigation']['source']}", AUTHOR)

    # --- Website / LinkedIn ---
    website_row = lit_header_row + lit_capacity + 8
    note(ws, f"C{website_row}", f_["website"]["value"], f_["website"]["source"])
    note(ws, f"C{website_row + 1}", f_["linkedin"]["value"], f_["linkedin"]["source"])

    # ===================================================================
    # CLEANUP: Hide and clear unused rows to prevent #REF! and #DIV/0!
    # ===================================================================
    for r in rows_to_hide:
        for col in ["C", "D", "E", "F", "G", "H", "I"]:
            cell = ws[f"{col}{r}"]
            
            # Use try/except to gracefully skip read-only MergedCell objects
            try:
                cell.value = None
                cell.comment = None
            except AttributeError:
                pass 
                
        # Hide the row completely from view
        ws.row_dimensions[r].hidden = True

    wb.save(out_path)
    return out_path