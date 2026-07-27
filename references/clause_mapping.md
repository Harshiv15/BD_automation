# Form 3CEB Part A — Clause-to-field mapping (Clauses 1-9)

Reference only — always verify against the current AY's actual form/utility, as wording and
clause numbering can shift year to year.

| Clause | Field | Typical source in financials |
|---|---|---|
| 1 | Name of assessee (note both old/new name if changed during the year) | Cover page / signing section |
| 2 | Address (foreign address if a foreign company) | Registered office note / cover page |
| 3 | PAN | Statutory audit report header, or ask user directly — often not in the financials at all |
| 4 | Nature of business/activity (code per ITR-6 instructions) | Directors' report / business overview section |
| 5-7 | [Verify against current form — often relate to accounting period / previous year] | — |
| 8 | Aggregate value of international transactions (per books of account) | Related-party transaction note, transactions with counterparties flagged as foreign/non-resident |
| 9 | Aggregate value of specified domestic transactions | Related-party transaction note, transactions with counterparties flagged as domestic related parties |

## Judgment calls that need a human, not this skill

- **Foreign vs domestic counterparty classification.** The financials often list related-party
  names without flagging residency status explicitly — this needs to be cross-checked against
  the group structure / AE list, not inferred from the transaction table alone.
- **AE relationship type** (subsidiary, fellow subsidiary, common directorship, etc. under
  Section 92A(2)) — surface the counterparty and value; the reviewer determines the
  relationship category.
- **Whether a related-party transaction is even in scope** — e.g. dividend transactions are
  typically excluded from the international-transaction aggregate even though they show up in
  AS-18/Ind AS 24 disclosures. Don't include these in the aggregate without reviewer
  confirmation.
