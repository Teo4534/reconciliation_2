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
python3 generate_fake_data.py examples --seed 20260101 >/dev/null
python3 build_ledger.py examples/roster.xlsx examples/bank.xlsx examples/fee_ledger.xlsx >/dev/null
python3 reconcile.py examples/fee_ledger.xlsx >/dev/null
python3 score.py | tail -9
WRONG=$(python3 score.py | awk '/^TOTAL/ {print $5}')
if [ "$WRONG" != "0" ]; then echo "FAIL: $WRONG misallocated receipt(s). Precision is the invariant."; exit 1; fi
echo "OK"
