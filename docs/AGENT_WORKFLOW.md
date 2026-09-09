# How changes to this repo were made

Most of the code here was written with Claude Code. This is the arrangement that was used, why it
was set up that way, and where a person decides.

## The gate: `check.sh`

One command that has to pass before anything is committed. The models do not decide whether a
change is correct; this does.

```
== 1/3 unit ==    tests/test_engine.py      21 tests, ~40ms   catches logic regressions
== 2/3 e2e  ==    tests/test_pipeline.py    31 tests, ~5s     three generated seeds, end to end
== 3/3 score ==   score.py, seed 20260101 in a temp dir       WRONG must be 0
```

The third gate is the one that matters: coverage may fall in a change, misallocations may not, and
the script exits non-zero if a single receipt is credited to the wrong family.
`.github/workflows/check.yml` runs the same script on every push and pull request, so the result is
visible to anyone reading the repository rather than only on my laptop.

The gate is what makes the rest of this work. Without a check that can say no, a writer and a
reviewer just agree with each other.

## The two roles: `.claude/agents/`

Two Markdown files. They are instructions, not programs: each tells a Claude Code session what it
is doing and which tools it may use.

**implementer** has edit tools. It reads `CLAUDE.md`, makes one change in `engine.py`, adds a unit
test in `tests/test_engine.py` that fails without the change, runs `./check.sh`, and reports what
changed, the test added and the last four lines of the gate.

**reviewer** has read-only tools and cannot edit. It sees only the diff and the gate output, never
the implementer's reasoning, re-runs `./check.sh` itself rather than trusting a pasted result,
tries to construct an input that credits the wrong family, and returns BLOCKING / SHOULD FIX /
NOTES with a verdict: APPROVE, FIX AND RE-REVIEW or REJECT.

The separation is the point. A model reviewing its own change is working from the reasoning that
produced it, so it re-finds the problems it already half-knew about. A reviewer starting cold from
the diff finds different ones. The failure this is aimed at is a change that lifts coverage by
quietly allocating a few receipts it should not: a writer pleased with the higher number will not
go looking for it, a reviewer told to break the change will.

## Where a person decides

Worth naming, because it is the part that is not automated.

- **The invariant is mine.** Never credit the wrong family, even at the cost of coverage. Nothing
  in this setup could have derived that; it comes from knowing what an arrears letter does to a
  parent who has already paid.
- **The gate encodes it.** Once written down it is enforced mechanically, on every change, by a
  script rather than by anyone's judgement in the moment.
- **I read both reports and decide.** The reviewer returns findings, not a verdict I have to
  accept. Some are wrong, and saying so is part of the loop.
- **Nothing merges without me.** Changes go on a branch and land through a pull request with the
  gate green on it.
- **Some things never go to a model.** The school's real roster and bank export are personal data
  about children and their parents. They are kept outside this repository, everything involving
  them is run locally, and no real name has ever been committed.

## The loop

Not configuration, just the instruction the session was given, capped at three rounds:

```
> Use the implementer agent to <task>. Then use the reviewer agent on the result. If the reviewer
> returns FIX AND RE-REVIEW, send its BLOCKING findings back to the implementer as the next task.
> Stop on APPROVE or after three rounds, and tell me which.
```

The cap matters: two models can disagree politely for a very long time about something genuinely
ambiguous, and at that point a person should look at it instead.

## Rules both roles work to

`CLAUDE.md`, in short: allocation logic lives in `engine.py` as pure functions; anything specific to
a term, year or school lives in `terms.py` or `Config`, never in the body of the code; every new
behaviour gets a unit test; tests are never edited to make them pass, and a test that looks wrong is
reported rather than changed; nothing is done until `./check.sh` is green with `WRONG = 0`.

## Known next tasks

1. Third-party payers in the fixture: grandparents, company accounts, a parent whose surname is not
   the family's. The real bank export has 23 of them and the generator produces none, so the tier
   ladder is currently measured where it matters least.
2. `rapidfuzz` in place of `difflib.SequenceMatcher` in `evidence()` - the only real hotspot at
   scale, and a true edit-distance ratio rather than a longest-common-subsequence one.
3. Cache `evidence()` per receipt - it is computed once for `named` in pass 1 and again in pass 2.
4. The reason string for an invoice number two families hold records only the reference, although
   the surname in the memo is what actually broke the tie. It understates the evidence used.
5. Aliasing across two families paying from one account (a grandparent), currently unsupported.
