# Running on real data

First real run: autumn 2026 term, Barclays export 15 Aug to 22 Sep 2026.
89 receipts. 75 allocated automatically (84%), 14 to review, 0 known wrong.

## What changed

The matching logic in `engine.py` (evidence, tiers, decisions) is untouched.
Everything below is loading, configuration, or a data quirk the generated fixtures never had.

### build_ledger.py

- **Config block.** Column aliases, term settings and filters now sit in one place at the top. A new term is a `TermSpec` entry; a new heading is one string in `ROSTER_COLS` or `BANK_COLS`.
- **`--term=ID`** picks the `TermSpec`. Default stays `JAN-2026`, so `check.sh` is unchanged.
- **Roster columns read by alias.** The real roster says `INVOICE N`, `REG FEE`, `SUPPLIES`, `ONEOFF REG`; the fixture says `INVOICE NUMBER`, `FEES`, `OFFICE SUPPLIES`, `REG FEES ONE OFF`. Both load. `PAID` is optional.
- **Note columns.** Previously any column starting `Unnamed: 1`. The real file has 16,359 empty columns Excel invented, so that rule grabbed thousands. Now: unlabelled columns that hold text.
- **Blank status rows.** The real roster leaves `Statut:` blank on rows added after the list was first drawn up (25 children, including two families with payments in the bank file). A blank status with an invoice number and a fee is now enrolled, with a note on the Children sheet.
- **Name order.** Late additions are typed `Chloe Renard`, first name first. The house convention is `RENARD Chloe`. When no token is capitalised, the last token is the surname (plus a lower-case particle before it). A middle name never joins it. If the row shares an invoice with a properly written row whose surname equals one of its tokens, it takes that surname (`Martin Lucas` on the MARTIN invoice is a MARTIN). Rows read this way get a note so the office can confirm.
- **Invoice number cleanup.** Leading apostrophe stripped (`'2026-999`). Stray letter in the year fixed (`2026t-999`). `?` means no invoice.
- **Name noise.** `FRATERIE`, `ESSAIS` and similar words the office types beside a name are dropped from the surname.
- **Unreadable bank lines stop the build.** A line with a memo but no parseable date or amount is reported, not dropped. Dates are read day first; amounts tolerate `£` and thousands separators.
- **Bank export read by heading.** Barclays writes a header row with `Number, Date, Account, Amount, Subcategory, Memo`. The fixture is headerless in a different order. When a header is present columns are taken by name. Legacy `.xls` needs `xlrd`.
- **Non-fee lines dropped.** The real statement is the whole account. Outgoing lines in `NOT_A_RECEIPT` categories (direct debits, card purchases, bill payments, bank interest) are removed. Only negative amounts qualify, so no receipt can be lost.
- **Out of sequence** is now judged against the term's invoice series, not a fixed `> 200`.

### engine.py

- `Config` gains `cur_series` and `prior_series`, inclusive bounds on the three-digit part of an invoice number.
- `parse_refs` only accepts a number for a term if it is inside that term's series as well as its year.

Why: autumn 2026 issues `2026-5xx` and `2026-6xx`. The prior term, January 2026, issued `2026-0xx` and `2026-1xx`. Same year, so the year alone cannot separate them. Both fields default to empty, which means "the year decides", the behaviour every existing fixture relies on.

### reconcile.py

- Position status says `in review`, not `unpaid`, for a family with nothing allocated but a receipt naming it in the Review queue.
- `config_for(wb)` reads the current and prior term, their invoice series and session counts off the ledger's own `Terms` sheet, and the per-session rates off `Rates`. No flag, and no way to reconcile a ledger against the wrong term.

## Review

Five adversarial reviewers (one lens each) plus one skeptic per finding, after the change. The real-data auditor re-checked all 74 automatic allocations against the raw roster and found none it could dispute, and no missing credit. Seven code findings survived and were fixed: a middle name could become surname evidence; any non-`Inscrit` status counted as blank; unreadable bank dates or amounts were dropped silently; rates and series were duplicated in `reconcile.py`; a `2026` literal in the Summary text.

### .gitignore

- Added `*.xls` and `*.xlsb`. The rules covered `*.xlsx` only. The Barclays export is `.xls` and would have been committed.

## Data quirks the first run surfaced

- Three parents typed an invoice number one off from their own (one digit off, three times). All three were held for review because the surname contradicted the reference.
- Seven invoice numbers are issued to two families. Each receipt against one of them was allocated only where the memo also named the family.
- One family is on the roster as `LI`. That is the real surname.
- `Moreau Lea Fontaine` is ambiguous under either name order. Rename it `MOREAU Lea Fontaine` on the roster.
- One payment from a family not on the term roster (`GARNIER`).

## How to run

```
pip install xlrd
python build_ledger.py roster.xlsx bank.xls ledger.xlsx --term=AUT-2026
python reconcile.py ledger.xlsx
```

Keep real files outside the repository or under an ignored name.
