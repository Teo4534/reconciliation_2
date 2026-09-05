# Fee reconciliation - working rules

Allocates bank receipts to families on ranked evidence. `engine.py` is the logic (pure functions,
no Excel); `reconcile.py` is the I/O around it. Read `README.md` for the tiers and why they exist.

## The one invariant
**A receipt is never allocated to the wrong family.** Coverage (the 86%) may go down in a change;
misallocations (the 0) may not. `score.py` reports `WRONG`; it must stay 0 on every seed.
Sending a line to Review is a cost. Crediting the wrong family is a wrong answer.

## Before you say a change is done
Run `./check.sh`. All three gates must pass. Paste the last four lines of its output in your report.
Do not edit tests to make them pass; if a test is wrong, say so and stop.

## Layout
- `engine.py`      allocation logic. Add behaviour here. Keep functions pure.
- `reconcile.py`   loads the ledger, calls the engine, writes Review / Position / Summary.
- `build_ledger.py` roster + bank export -> structured workbook.
- `score.py`       accuracy against ground truth, by failure mode.
- `tests/test_engine.py`   fast unit tests (ms). Add one for every new behaviour.
- `tests/test_pipeline.py` end-to-end over three generated seeds.

## Style
Small pure functions with a docstring saying what they decide. No hard-coded terms, years or fee
amounts outside `Config`. Reason strings are templates: the same input always produces the same text.

## Workflow
Changes go implementer -> harness -> reviewer -> (fix if needed) -> done.
Use the `implementer` agent to make a change and the `reviewer` agent to check it. The reviewer sees
only the diff and the harness output, never the implementer's reasoning - that separation is the point.
