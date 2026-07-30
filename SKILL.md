---
name: tp-bd-factsheet-filler
description: Extracts company information from an Annual Report (AR)/financial statements PDF - and, when the AR doesn't have it, the company's website/LinkedIn - to populate a TP Business Development Fact Sheet Excel template (headquarters, business description, standalone/consolidated financials, AE revenue split, shareholding, related-party transactions, litigation, website/LinkedIn). Use whenever the user provides an AR/financial statements and a BD Fact Sheet template and asks to fill, draft, or prepare a BD, fact sheet, or company profile for TP purposes. Always flag extracted figures for human review - never treat output as final.
compatibility: Requires Python 3 with pdfplumber and openpyxl installed.
---

# TP BD Fact Sheet Filler

## Scope and non-negotiables

Every output is a **draft for human review**. Attach a source-citing comment to every filled
cell - reviewers rely on these to spot-check quickly, so never skip them.

**Never touch cell formatting.** Only write cell values and comments. Confidence is communicated only through the comment text now (e.g. "high confidence - directly stated in the AR" / "flagged - judgment call, see note").

**Never reference other companies' fact sheets anywhere in the output.** Each BD is a completely isolated description of one company.

## Extraction preference: transcribe, don't summarize

For factual/tabular fields (litigation entries, shareholder names, statute names, related-party
transaction labels, countries), **transcribe as close to verbatim from the source as possible**.

## Field-by-field guidance

### Headquarters / group info
- India entity HQ: registered office address, verbatim, from the AR.
- **Group HQ / group description**: first check whether the company has a holding company at all. If there's genuinely no group, write **`N/A`**.

### Company / group description
Short, factual, paraphrased (2-4 sentences): what the company does, listing status, segment structure. 

### Standalone / consolidated summary of operations
- **Turnover = Revenue from Operations ONLY.** DO NOT include Other Income. 
- **Total cost = Total Expenses.**
- **Note:** Because you are excluding Other Income from Turnover, the template's calculated PBT (Turnover - Total Cost) will likely not match the AR's reported PBT. This is expected.
- **Consolidated block**: if the company has no subsidiaries/associates/JVs, write `N/A`.

### AE / Domestic / Export revenue split
- Base the split on **Revenue from Operations**.
- Check whether related parties are foreign (real AEs) or domestic.
- **AE revenues (Sr 9)** MUST specifically be the total revenue from the **sale of services** to **non-Indian (foreign) AEs**, as disclosed in the RPT section.

### Shareholding
- **List only DIRECTOR shareholders individually**. Group every other promoter-family member who isn't a director into one **"Other Promoters (non-director)"** row.
- **"Others (<5% equity per shareholder)"** - non-promoter holders individually below 5% each.
- **"Others (not listed in AR)"** - a balancing plug.

### Related party transactions (CRITICAL RULES)
- **EXCLUDE INDIVIDUALS:** You must completely ignore any transaction where the related party is a natural person/individual (e.g., Directors, Promoters, Key Managerial Personnel). **Only extract transactions with corporate entities (Companies, LLPs, Trusts).**
- **COLLATE BY NATURE:** If there are multiple transactions of the exact same nature (e.g., "Business Support Services" provided to three different companies), **collate them together and sum their values into one single row**. 
- **VERBATIM LABELS:** Use the EXACT transaction labels exactly as they appear in the AR's Related Party table. **DO NOT** append the company name or entity name to the label (e.g., write "Business Support Services", never "Business Support Services - Tips Films Limited").
- **Sign convention:** positive = amount payable BY the company TO the related party (a cost/outflow); negative = a net amount RECEIVABLE BY the company FROM the related party.

### Litigation (CRITICAL RULES)
- **ONLY FROM AR:** You must extract litigation details *exclusively* from the "Contingent Liabilities" or "Litigations" section of the Annual Report PDF. 
- **NO HALLUCINATION/WEB SEARCH:** Do not use external knowledge or hallucinate penalty amounts. If the AR does not explicitly list the "Amount paid" or "Forum where dispute is pending", output `null` or `"-"`.
- Transcribe the statute name, nature of dues, amounts demanded/paid, period, and forum exactly as tabulated in the AR.

### Countries where they have presence
- **Country names only**. Base this on where the entity and its disclosed related parties/subsidiaries are actually incorporated.

### Website / LinkedIn
- Website: check the AR first. LinkedIn: search the web if possible, otherwise `null`.

## Workflow (For Downstream Systems)
1. Extract structured data from the AR PDF natively.
2. Auto-detect the amount unit stated in the financial statements header (Rs Lakhs / Rs Millions / Rs Crores) and convert everything to lakhs.
3. Output the strict JSON schema required.