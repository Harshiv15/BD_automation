"""
extract_financials.py

Extracts candidate Form 3CEB Part A inputs from a company's financial statements PDF.

This is a STARTING POINT, not a finished parser. Real annual report PDFs vary hugely in
layout — you will need to tune the table-detection logic (or add a per-client config) once
you see real (anonymized) samples.

Usage:
    python extract_financials.py --pdf financials.pdf --out extracted.json
"""

import argparse
import json
import re
import sys

try:
    import pdfplumber
except ImportError:
    sys.exit("Missing dependency: pip install pdfplumber --break-system-packages")


PAN_PATTERN = re.compile(r"\b[A-Z]{5}[0-9]{4}[A-Z]\b")
RELATED_PARTY_HEADERS = [
    "related party", "related parties", "associated enterprise",
    "as-18", "ind as 24", "transactions with related part",
]


def find_related_party_section(pages_text):
    """Return page indices whose text suggests a related-party disclosure note."""
    hits = []
    for i, text in enumerate(pages_text):
        lowered = text.lower()
        if any(h in lowered for h in RELATED_PARTY_HEADERS):
            hits.append(i)
    return hits


def extract(pdf_path):
    result = {
        "assessee_name": None,
        "pan": None,
        "candidate_related_party_pages": [],
        "related_party_tables": [],  # list of {page, table_rows}
        "notes": [],
    }

    with pdfplumber.open(pdf_path) as pdf:
        pages_text = [p.extract_text() or "" for p in pdf.pages]

        # PAN — simple regex scan; verify manually, false positives are possible
        for text in pages_text:
            m = PAN_PATTERN.search(text)
            if m:
                result["pan"] = m.group(0)
                break

        # Related-party / AE disclosure pages
        rp_pages = find_related_party_section(pages_text)
        result["candidate_related_party_pages"] = [p + 1 for p in rp_pages]  # 1-indexed for humans

        for page_idx in rp_pages:
            page = pdf.pages[page_idx]
            tables = page.extract_tables()
            for t in tables:
                result["related_party_tables"].append({
                    "page": page_idx + 1,
                    "rows": t,
                })

        if not rp_pages:
            result["notes"].append(
                "No related-party/AE disclosure section detected by keyword scan — "
                "search manually, this financial statement may use different terminology."
            )

    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    data = extract(args.pdf)

    with open(args.out, "w") as f:
        json.dump(data, f, indent=2)

    print(f"Wrote {args.out}. Review candidate_related_party_pages and related_party_tables "
          f"manually before using in the 3CEB filler — table extraction from PDFs is not "
          f"reliable enough to trust unreviewed.")


if __name__ == "__main__":
    main()
