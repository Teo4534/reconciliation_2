# How this repo is developed with agents

## The gate: `check.sh`

The agents do not decide whether a change is correct; `check.sh` does. The third gate is the one
that matters: coverage may fall in a change, misallocations may not.

```
== 1/3 unit ==    tests/test_engine.py      21 tests, ~40ms   catches logic regressions
== 2/3 e2e  ==    tests/test_pipeline.py    30 tests, ~5s     three generated seeds, end to end
== 3/3 score ==   score.py, seed 20260101 in a temp dir       WRONG must be 0
```

## The agents: `.claude/agents/`

**implementer** has edit tools. It reads `CLAUDE.md`, makes one change in `engine.py`, adds a unit
test in `tests/test_engine.py` that fails without the change, runs `./check.sh`, and reports what
changed, the test added and the last four lines of the gate.

**reviewer** has read-only tools and cannot edit. It sees only the diff and the gate output, re-runs
`./check.sh` itself, tries to construct an input that credits the wrong family, and returns
BLOCKING / SHOULD FIX / NOTES with a verdict: APPROVE, FIX AND RE-REVIEW or REJECT.

Both work to the rules in `CLAUDE.md`: logic lives in `engine.py` as pure functions; anything term-,
year- or school-specific lives in `Config`; every new behaviour gets a unit test; tests are never
edited to pass; and nothing is done until `./check.sh` is green with `WRONG = 0`.

## The loop

The main Claude Code session does no editing itself; it runs this prompt, capped at three rounds:

```
> Use the implementer agent to <task>. Then use the reviewer agent on the result. If the reviewer
> returns FIX AND RE-REVIEW, send its BLOCKING findings back to the implementer as the next task.
> Stop on APPROVE or after three rounds, and tell me which.
```

## Known next tasks

1. `rapidfuzz` in place of `difflib.SequenceMatcher` in `evidence()` - the only real hotspot at
   scale, and a true edit-distance ratio rather than a longest-common-subsequence one.
2. Cache `evidence()` per receipt - it is computed once for `named` in pass 1 and again in pass 2.
3. Amounts as weak evidence for tie-breaks between same-surname families, never the sole basis.
4. Aliasing across two families paying from one account (a grandparent), currently unsupported.
