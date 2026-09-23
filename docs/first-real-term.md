# First real term: what happened when the engine met the bank

Until this week every number in this repository came from `generate_fake_data.py`. The engine allocated 86% of a fictional term's receipts with none wrong, and the harness proved it on three seeds. Then the real autumn 2026 roster and a Barclays export arrived. This is what that looked like.

## The engine was fine. The files were not.

Not one matching rule needed to change. Everything that broke was between the files and the engine.

- The bank export was a legacy `.xls`, with a header row and the columns in a different order from the fixture.
- The roster called its columns `INVOICE N` and `REG FEE`, not `INVOICE NUMBER` and `FEES`, and had 16,359 empty columns Excel had invented. The old note-column rule would have grabbed thousands of them.
- 25 children had a blank status. They were late joiners, typed in below the main list, first name first instead of `SURNAME First`. Two of them had payments in the bank file. The old filter dropped all 25.
- Invoice numbers had leading apostrophes, a stray `t` in the year, and `?` where none had been issued.
- The statement was the whole account: rent, utilities, a 5p interest charge, alongside the fee receipts.
- This term's invoices are `2026-5xx`. Last term's were `2026-0xx`. Same year, so the year could no longer tell a current reference from a prior one.

Lesson one: the hard part of a matching system is rarely the matching. It is the ten small ways the input differs from what you imagined.

## What changed

All loading and configuration. `docs/real-data.md` has the detail.

- A config block at the top of `build_ledger.py`: column aliases and filters. A new heading is one string. Terms stay in `terms.py`, one entry per term.
- Bank file read by heading when it has one. Outgoing non-fee lines dropped, negatives only, so no receipt can be lost.
- Blank status with an invoice and a fee counts as enrolled, with a note for the office.
- `Config` in `engine.py` gained optional invoice-series bounds so two terms in one year stay apart. Empty by default; every existing fixture behaves as before.
- `.gitignore` now covers `*.xls`. It did not, and the bank file would have gone into a public repo.

The harness output was identical before and after: 60/77 allocated on the fixture, none wrong.

## The result

89 receipts. 75 allocated automatically (84%). 14 to review. £23,370 of £27,817 placed.

The three most useful lines were refusals. Three parents typed an invoice number one off from their own (one digit off, three times). The engine held all three because the surname in the memo contradicted the reference. A VLOOKUP credits the wrong family three times.

Seven invoice numbers had been issued to two families. Each payment against one was allocated only where the memo also named the family.

## The review

After the change, five reviewers each took one lens (engine, roster loading, bank loading, reconcile config, and an independent audit of the real allocations), and every finding went to a skeptic told to refute it. 15 raw findings, 7 survived.

The audit could not dispute any of the 74 allocations and found no missing credit.

The worst code finding was mine. The new name rule took everything after the first token as the surname, so a roster row `Grace Bayo Vasseur` would have made BAYO surname evidence for a child who is not a Bayo. The reviewer reproduced it end to end with a three-row roster and a one-line bank file: an automatic credit to the wrong family, where the old code would have sent it to review. Fixed: the last token is the surname, and a middle name never joins it. Fixing it also freed one real review line, a LEROUX payment that had been held because a mis-parsed "Lea" looked like a surname. A second rule followed from the office's own data: when an inferred surname shares an invoice with a properly written row, and one of its tokens equals that row's surname, it adopts it. `Vincent Bernard` on 2026-999 is a third VINCENT child, not a BERNARD family.

The Position sheet also stopped saying "unpaid" for a family whose payment is sitting in Review. It says "in review". The first is a chase; the second is a job for the office; conflating them is how a parent who has paid gets a reminder.

Lesson two: a bug in code can be reproduced with three rows and a test. A bug in a judgement call is a wrong number in a cell that nobody finds.

## On the AI

Most of the code in this change was written with Claude. The invariant, the failure-mode design, the harness, the two design decisions (adapt the repo rather than write throwaway glue; treat late joiners as enrolled) and the verification were the work. The AI's one real mistake was found because it was in code, by a review the code made possible.

## The mistake that mattered

The README says no real data appears in this repository. While writing this post and the changelog, I put real family surnames into the docs, into code comments, into one test, and into a commit message, and pushed the branch. They were public for about thirty minutes.

The tooling did not catch it. The harness checks allocations, not prose. `.gitignore` blocks the spreadsheets, not a name typed into a Markdown file. What caught it was reading the branch on GitHub and asking why a child's surname was in a commit message.

What followed: the branch was deleted from GitHub within minutes; every commit on it was rebuilt from scratch with the names replaced by fictional ones, in files and in messages; the full history was searched for every name before the branch went back up; GitHub Support was asked to purge the unreachable commits. `main` was never touched.

Two rules came out of it. Nothing derived from a real file goes into any tracked file or commit message, ever: examples use invented names, always. And the person is the last gate, not the first; the tool that produced the leak was the same one that produced the fix, and neither happened without someone reading the output.

Lesson three: the invariant covers the ledger. It did not cover the write-up. Now it does.

## What the office does next

Work the 14 review lines. Rename `Moreau Lea Fontaine` to `MOREAU Lea Fontaine`. Add LEGRAND to the roster; a payment of £220 arrived and there is no row to put it against. Re-run in two weeks with a fresh export; 40 families show unpaid three weeks into term, another 13 are marked "in review" because a payment naming them is waiting on a decision, and 8 only just joined, so that run is the one that says whether anyone actually left.
