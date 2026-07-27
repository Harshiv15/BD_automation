---
name: tp-bd-factsheet-filler
description: Extracts company information from an Annual Report (AR)/financial statements PDF - and, when the AR doesn't have it, the company's website/LinkedIn - to populate a TP Business Development Fact Sheet Excel template (headquarters, business description, standalone/consolidated financials, AE revenue split, shareholding, related-party transactions, litigation, website/LinkedIn). Use whenever the user provides an AR/financial statements and a BD Fact Sheet template and asks to fill, draft, or prepare a BD, fact sheet, or company profile for TP purposes. Always flag extracted figures for human review - never treat output as final.
compatibility: Requires Python 3 with pdfplumber and openpyxl installed.
---

# TP BD Fact Sheet Filler

## Scope and non-negotiables

Every output is a **draft for human review**. Attach a source-citing comment to every filled
cell - reviewers rely on these to spot-check quickly, so never skip them.

**Never touch cell formatting.** Only write cell values and comments - never fill color,
font, border, alignment, or size. An earlier version of this skill applied background-color
highlighting per confidence level; that changed the template's appearance and was explicitly
rejected. Confidence is communicated only through the comment text now (e.g. "high confidence -
directly stated in the AR" / "flagged - judgment call, see note").

**Never reference other companies' fact sheets anywhere in the output.** Each BD is a
completely isolated description of one company - no comparisons to, or mentions of, any other
client's BD in values or comments.

## Extraction preference: transcribe, don't summarize

For factual/tabular fields (litigation entries, shareholder names, statute names, related-party
transaction labels, countries), **transcribe as close to verbatim from the source as possible**
rather than summarizing - exact transcriptions are faster for a human to review against the
source than a paraphrase is. Reserve paraphrasing for genuinely prose fields (e.g. the company
description), and even there keep it short (2-4 sentences).

## Field-by-field guidance

### Headquarters / group info
- India entity HQ: registered office address, verbatim, from the AR.
- **Group HQ / group description**: first check whether the company has a holding company at
  all (search for "does not have a holding company" / "doesn't have any
  holding/subsidiary/associate companies", or check the shareholding pattern for a corporate
  vs. individual/family promoter). If there's genuinely no group, write **`N/A`** - keep any
  explanation in the comment, not the cell value. If there IS a group but the exact HQ isn't
  stated in the AR, check the web for the parent entity before settling for a partial answer.

### Company / group description
Short, factual, paraphrased (2-4 sentences): what the company does, listing status, segment
structure. Pull from an "About"/entity-background note or MD&A section; check the website next
if not in the AR.

### Standalone / consolidated summary of operations
- **Turnover = Total Income** (Revenue from Operations + Other Income), and **Total cost =
  Total Expenses**, both as literally reported. **Leave PBT and Margin as the template's own
  formulas** (Turnover - Total Cost, and PBT/Total Cost) rather than hardcoding, whenever that
  formula reconciles to the AR's reported PBT (it usually will, since Total Income already
  includes Other Income). Matching the reported PBT via the formula matters more than any
  other section of the sheet (e.g. the AE/export split, which is based on Revenue from
  Operations alone) summing back to this Turnover figure exactly - that mismatch is expected
  and fine; note it in a comment rather than trying to force reconciliation there too.
  - Only if the AR has an exceptional item or similar that breaks the Turnover - Total Cost =
    PBT identity even after using Total Income, hardcode the literal reported PBT figure
    instead and flag why in a comment.
- **Consolidated block**: if the company has no subsidiaries/associates/JVs, overwrite the
  block's formula cells with literal `N/A` text (a value change, not a formatting change - fine
  under the no-formatting-changes rule).

### AE / Domestic / Export revenue split
- Base the split on **Revenue from Operations** (not Total Income) - e.g. apply an
  export-contribution disclosure ("contribution of exports is X% of total turnover", often in a
  BRSR "markets served" section) to Revenue from Operations to get domestic vs. export.
- Check whether related parties are foreign (real AEs) or domestic - revenue from foreign
  THIRD-PARTY customers (e.g. licensing to foreign platforms) is export revenue, not AE revenue,
  unless that customer is an actual disclosed related party.
- **"AE revenues" (Sr 9) should be cross-checked against the Related Party Transactions
  table**: sum whichever RPT line items represent revenue earned BY the company from a related
  party. If none exist, AE revenues = 0 - say explicitly in the comment that this was
  cross-checked against the RPT table, not just assumed absent.

### Shareholding
- **List only DIRECTOR shareholders individually** (cross-check the Board of
  Directors/Directors' Report to see which named promoters actually hold a directorship).
  Group every other promoter-family member who isn't a director into one **"Other Promoters
  (non-director)"** row.
- Use exactly these labels, consistently, so different BDs read the same way:
  - **"Other Promoters (non-director)"** - promoter-group members who aren't directors.
  - **"Others (<5% equity per shareholder)"** - non-promoter holders individually below 5%
    each (fragmented public/institutional float).
  - **"Others (not listed in AR)"** - a balancing plug (use a formula:
    `=$Total - SUM(named rows)`) when the AR only discloses a partial list (e.g. top 10) and
    the remainder can't be attributed to any named/banded category.
- Both FY columns' percentages should sum to ~100% - check this in the proofreading pass, not
  just by eye. Total share counts can differ year to year (buybacks) - don't assume they match.
- If more rows are needed than the template currently has, the downstream Python logic will use `safe_insert_rows.py` dynamically to add rows IN THE MIDDLE of the table (e.g., if there are 3 new shareholders, insert 3 rows *before* the final "Total" row). Do NOT append rows after the "Total" row.

### Related party transactions (CRITICAL RULES)
- **EXCLUDE INDIVIDUALS:** You must completely ignore any transaction where the related party is a natural person/individual (e.g., Directors, Promoters, Key Managerial Personnel). Only extract transactions with corporate entities (Companies, LLPs, Trusts).
- **VERBATIM LABELS:** Use the exact transaction labels exactly as they appear in the AR's Related Party table (e.g., "Purchase of Assets", "Reimbursement of Expenses"). **DO NOT** append the company name or entity name to the label (e.g., never write "Purchase of Assets - Tips Films Limited" if the AR just says "Purchase of Assets"). Keep the label generic as printed.
- **Sign convention** (state this in a comment wherever a negative value appears): positive =
  amount payable BY the company TO the related party (a cost/outflow); negative = a net amount
  RECEIVABLE BY the company FROM the related party (e.g. a reimbursement due back). Apply this
  consistently across every RPT row.
- Keep this table to genuine inter-entity commercial transactions - exclude KMP-level personal
  compensation (director remuneration, salary, sitting fees) and exclude balance-sheet items
  (outstanding payables/receivables belong elsewhere in the sheet). Note what was excluded and
  why in a comment.

### Litigation (CRITICAL RULES)
- **ONLY FROM AR:** You must extract litigation details *exclusively* from the "Contingent Liabilities" or "Litigations" section of the Annual Report PDF. 
- **NO HALLUCINATION/WEB SEARCH:** Do not use web search, external knowledge, or hallucinate penalty amounts. If the AR does not list the "Amount paid" or "Forum where dispute is pending", output `null` or `"-"`.
- **List everything material, not just TP-specific disputes** - tax, GST/service tax, FEMA, or
  any other statutory dispute in the contingent liabilities note. Transcribe the statute name,
  nature of dues, amounts demanded/paid, period, and forum exactly as tabulated.

### Countries where they have presence
- **Country names only** - nothing else, no descriptive prose about markets served or
  distribution reach. Base this on where the entity and its disclosed related
  parties/subsidiaries are actually incorporated, not where products are distributed.

### Website / LinkedIn
- Website: check the AR first. LinkedIn: search the web (virtually never in the AR).

## Workflow (For Downstream Systems)
1. Extract structured data from the AR PDF natively.
2. Auto-detect the amount unit stated in the financial statements header (Rs Lakhs / Rs Millions / Rs Crores) and convert everything to lakhs.
3. Output the strict JSON schema required.
4. Downstream Python will execute `safe_insert_rows.py` based on the JSON length and populate the Excel template.