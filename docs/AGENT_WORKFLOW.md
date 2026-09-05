# How this repo is developed with agents

This is the workflow behind the "How changes are made" section of the README, in enough detail to
run it. It is deliberately small: one gate, two agents, one loop.

## The gate: `check.sh`

Agents do not test code. The harness tests code; agents loop around it. Without a deterministic
check in the middle, a writer and a reviewer just agree with each other, because the reviewer has
no ground truth to disagree from.

```
./check.sh
== 1/3 unit ==    tests/test_engine.py      21 tests, ~40ms   catches logic regressions
== 2/3 e2e  ==    tests/test_pipeline.py    30 tests, ~5s     three generated seeds, end to end
== 3/3 score ==   score.py on examples/     WRONG must be 0   the invariant this repo exists for
```

The third gate is the one that matters. Coverage (the 86% allocated automatically) may fall in a
change. Misallocations may not. A receipt credited to the wrong family becomes an arrears letter to
a parent who has paid.

## The agents: `.claude/agents/`

Both are Markdown files with YAML frontmatter. The frontmatter says when to use the agent, which
tools it may use, and which model runs it. The body is its system prompt.

**implementer** - has edit tools. Gets one task, reads `CLAUDE.md`, changes `engine.py`, adds a unit
test in `tests/test_engine.py` that fails without the change, runs `./check.sh`, and reports in a
fixed shape: what changed, the test added, the last four lines of the gate, anything uncertain.

**reviewer** - read-only tools. Cannot edit. Receives only the diff and the gate output, never the
implementer's reasoning, and is told to construct an input that misallocates. It must re-run
`./check.sh` itself rather than trust a pasted result. It returns BLOCKING / SHOULD FIX / NOTES and
a verdict.

The separation is the point. A model reviewing its own change is conditioned on the reasoning that
produced it, so it finds the flaws it was already half-aware of. A reviewer starting cold from the
diff finds the others. The failure mode that matters here is a change that lifts coverage from 86%
to 90% by quietly allocating three receipts it should not; a writer proud of the 90% will not see
it, a reviewer told to break it will.

## The loop

The loop is not configuration. It is a prompt the main session follows:

```
cd reconciliation-repo && claude

> Use the implementer agent to replace SequenceMatcher in engine.evidence() with rapidfuzz,
> keeping the 0.85 threshold semantics as close as possible. Then use the reviewer agent on
> the result. If the reviewer returns FIX AND RE-REVIEW, send its BLOCKING findings back to
> the implementer as the next task. Stop on APPROVE or after three rounds, and tell me which.
```

That is an orchestrator (the main session, which does no work itself), two sub-agents, a hand-off
contract (the fixed report shapes), a harness, and a bounded loop. The round cap matters: two agents
can ping-pong forever on a genuinely ambiguous point.

`/agents` inside Claude Code lists both agents under the project scope, which is how you know the
files are being picked up.

## Working in parallel

Independent tasks can run as separate sessions on separate git worktrees, so nothing collides:

```
git worktree add ../recon-fuzz -b rapidfuzz        # swap the fuzzy matcher
git worktree add ../recon-perf -b evidence-cache   # stop computing evidence() twice per receipt
git worktree add ../recon-docs -b readme           # rewrite the README
```

One `claude` per directory, one task each, `./check.sh` before reporting. Review three diffs, merge
the ones that pass. Worktrees isolate the filesystem the way sub-agents isolate context.

## Scheduling

Once the gate is trusted, the loop can run unattended with `/loop`:

```
/loop every 2 hours: run ./check.sh on main. If any gate fails, use the implementer agent to fix
it and the reviewer agent to check the fix. Open a PR with the reviewer's report as the
description. Never merge.
```

The loop opens the PR. A person merges it.

## Rules both agents work to

See `CLAUDE.md`. In short: logic lives in `engine.py` as pure functions; anything term-, year- or
school-specific lives in `Config`; every new behaviour gets a unit test; tests are never edited to
pass; and nothing is done until `./check.sh` is green with `WRONG = 0`.

## Known next tasks

Good first runs for the loop, in rough order of value:

1. `rapidfuzz` in place of `difflib.SequenceMatcher` in `evidence()` - the only real hotspot at
   scale, and a true edit-distance ratio rather than `SequenceMatcher`'s longest-common-subsequence
   one.
2. Cache `evidence()` per receipt - it is computed once for `named` in pass 1 and again in pass 2.
3. Amounts as weak evidence for tie-breaks between families with the same surname, never as the
   sole basis for allocation.
4. Aliasing across two families paying from one account (a grandparent), currently unsupported.
