# Fee reconciliation

A small charity invoices termly school fees and receives them by bank transfer. The bank memo is
the only link between a payment and an invoice, and most of the time that link is broken: parents
mistype the reference, reuse last term's, run it together with their name, or send nothing but a
surname. Families with several children pay one combined invoice, often in instalments, sometimes
from an account in a different name.

The result is that nobody can answer "who still owes us money?" without a day of manual work, so
it gets done once a term and part-payments go unnoticed until year end.

This repository builds a fee ledger from the roster and the bank export, then allocates each
receipt to a family on ranked evidence, allocating what it can prove and queuing the rest for a
human with the reasoning written out.

**On the sample data in `examples/`: 86% of receipts are allocated with no human involvement, and
none is allocated to the wrong family.** The remaining 11 go to a review queue.

![The Position sheet](examples/position_sheet.png)

No real data appears anywhere in this repository. `generate_fake_data.py` produces a fictional
roster and bank export that reproduce the failure patterns, not the records, which also means
every line has a known correct answer, so accuracy can be measured rather than asserted.

## Run it

```bash
pip install -r requirements.txt

python generate_fake_data.py examples --seed 20260101      # roster.xlsx, bank.xlsx, ground_truth.csv
python build_ledger.py examples/roster.xlsx examples/bank.xlsx examples/fee_ledger.xlsx
python reconcile.py examples/fee_ledger.xlsx               # allocates, writes Position / Review / Summary
python score.py                                            # measures the result against ground truth
pytest -q                                                  # 10 tests across 3 generated datasets
```

## What the bank actually sends

The memo is a fixed-width field: 23 characters of payer name, then a 21-character reference slot
that silently truncates. Every row below is real output from `examples/`, with the tier the engine
assigned and the reason it recorded.

| what went wrong | bank memo | tier | engine reason |
|---|---|---|---|
| nothing | `MARGIT WENDELIN     2026-047` | A | invoice reference 2026-047 |
| glued to the name | `GENEVIEVE MAALOUF   MAALOU2026-030` | A | invoice reference 2026-030 |
| hyphen dropped | `VERA FALZON         2026018` | A | invoice reference 2026-018 |
| only the number | `ELODIE HASANOVIC    060` | B | surname HASANOVIC |
| cut off by the bank | `ULRICH CARDOSO      CARDOSO FRENCH SCHOOL` | B | surname CARDOSO |
| no reference at all | `ULRICH QUILLIAM     school fees` | B | surname QUILLIAM |
| a child's name only | `ZORA VANTERPOOL     Halima` | B | surname VANTERPOOL |
| second instalment | `GENEVIEVE MAALOUF   2nd payment` | A | payer name previously seen with a verified reference for this family |
| last term's reference | `MARGIT PELLETIER    2025-627` | D | several families fit: FAM-042 (surname PELLETIER); FAM-045 (surname PELLETIER) |
| **someone else's reference** | `MARGIT RASMUSSEN    2026-039` | X | reference 2026-039 points to FAM-049 but the memo names FAM-048, parent may have typed the wrong invoice number |

The last row is the case that matters. A naive matcher follows the reference, credits the wrong
family, and produces two errors at once: one parent chased for money they paid, another marked
paid when they haven't. The engine refuses to allocate when the reference and the payer name
disagree, and says so.

## How a receipt gets a family

Ranked evidence, first hit wins.

| tier | evidence | allocated? |
|---|---|---|
| M | a person typed a family ID into the override column | yes |
| A | current-term invoice reference naming exactly one family, uncontradicted by the memo, or a payer name already seen with a verified reference | yes |
| B | exactly one roster surname appears in the memo | yes |
| C | fuzzy evidence: a truncated or misspelt surname, or a child's first name | review |
| D | nothing usable, or several families fit | review |
| X | the reference and the payer name point to different families | review |

Two things do most of the work. **Aliasing**: confirm a payer once, by reference or by hand, and
every later payment from that account follows, including ones with no reference at all, which is
how second instalments get matched. **Refusing to guess**: tiers C, D and X are held back with
their candidates and reasoning, so a person spends their time on the 14% that need judgement.

Every reason string is a template, not a generated sentence. Each rule that fires appends a fixed
phrase. The same input always produces the same allocation and the same explanation, which is what
makes the output auditable.

## Measured on generated data

`python score.py`, seed 20260101:

```
failure mode        lines   auto  correct  WRONG  review
clean                  17     17       17      0       0
prior_term             14     12       12      0       2
name_only              11      9        9      0       2
glued                   9      9        9      0       0
no_hyphen               7      7        7      0       0
truncated               6      4        4      0       2
bare                    5      5        5      0       0
first_name_only         4      4        4      0       0
wrong_ref               4      1        1      0       3
shared_invoice          2      1        1      0       1
outflow                 1      0        0      0       1
TOTAL                  80     69       69      0      11

allocated automatically : 69/80 = 86%
of those, correct       : 69/69 = 100.0%
sent for human review   : 11
```

The design target is precision, not coverage. Sending more lines to a human is a cost; crediting
one to the wrong family is a wrong answer that propagates into arrears letters. `pytest` asserts
zero misallocations across three independently generated datasets.

## The ledger

`build_ledger.py` writes a workbook rather than a database, because the people who use it live in
spreadsheets and need to see what happened.

- **Position**: one row per family: expected, received, balance, status, every bank reference they
  used, and a plain-English flag such as *£520.50 tagged AUT-2025 looks like a JAN-2026 payment*.
  Green settled, red owing, amber check before chasing.
- **Review**: the only sheet needing a human. Decide, type the family ID into `Receipts` column F,
  re-run `reconcile.py`.
- **Receipts**: every bank line and what the engine did with it.
- **Children / Families**: the roster, with each child's fee computed from the rules and compared
  against what was actually invoiced, so mis-invoiced pupils surface automatically.
- **Rates / Terms**: the fee rules as data. A price change is a cell edit; nothing is hard-coded.

Two structural decisions worth explaining, because they came out of the real data:

**The invoice number cannot be the family key.** Unrelated families were issued the same number,
and numbers change every term, so history would be lost. Families get a stable `FAM-nnn` key and
invoice numbers hang off it.

**Fee rules are data, not code.** Sibling tiers, session counts, registration and supplies charges
live in `Rates` and `Terms`. Running the rules against the invoices as a check found a pupil billed
£18 for supplies instead of £25, an error nobody had spotted.

## Limitations

- First-name matching only works while a first name is unique on the roster; a second child with
  the same name silently drops it to tier D.
- Aliasing assumes one bank account belongs to one family. It held on the data tested, but a
  grandparent paying for two families would break it.
- Ageing runs from the term start, not from invoice date, because invoice dates were not available.
- A term whose invoice register has not been loaded can have its receipts identified by family but
  not reconciled against an invoice; those are flagged, not silently ignored.
- Amounts are not used as matching evidence, only as a sanity check shown to the reviewer.

## Files

```
generate_fake_data.py   fictional roster + bank export + ground truth
build_ledger.py         roster + bank  ->  structured workbook
reconcile.py            allocation engine, Position / Review / Summary sheets
score.py                accuracy against ground truth, by failure mode
tests/test_pipeline.py  10 end-to-end tests over 3 generated datasets
examples/               generated data and the resulting ledger
```

MIT licensed.
