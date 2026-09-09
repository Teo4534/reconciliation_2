#!/usr/bin/env bash
# The harness. One command, three gates. Agents and humans run exactly this.
#   1. unit tests on the engine        (milliseconds - catches logic regressions)
#   2. end-to-end on three seeds       (seconds     - catches integration breaks)
#   3. precision on the example data   (must be 0 WRONG - the invariant this project exists for)
set -euo pipefail
cd "$(dirname "$0")"
echo "== 1/3 unit =="; python3 -m pytest tests/test_engine.py -q
echo "== 2/3 e2e  =="; python3 -m pytest tests/test_pipeline.py -q
echo "== 3/3 score =="
# Gate 3 runs in a temp dir; examples/ is no longer touched (openpyxl restamps timestamps, dirtying git on every run).
# To refresh examples/ by hand, run the same three commands below with examples/ in place of $TMP.
TMP=$(mktemp -d); trap 'rm -rf "$TMP"' EXIT
python3 generate_fake_data.py "$TMP" --seed 20260101 >/dev/null
python3 build_ledger.py "$TMP/roster.xlsx" "$TMP/bank.xlsx" "$TMP/fee_ledger.xlsx" >/dev/null
python3 reconcile.py "$TMP/fee_ledger.xlsx" >/dev/null
SCORE=$(python3 score.py "$TMP/fee_ledger.xlsx" "$TMP/ground_truth.csv")
printf '%s\n' "$SCORE" | tail -9
WRONG=$(printf '%s\n' "$SCORE" | awk '/^TOTAL/ {print $5}')
if [ "$WRONG" != "0" ]; then echo "FAIL: $WRONG misallocated receipt(s). Precision is the invariant."; exit 1; fi
echo "OK"
