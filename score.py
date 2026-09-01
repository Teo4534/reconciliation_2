"""
score.py — measure the engine against the generator's ground truth.

    python score.py [ledger.xlsx] [ground_truth.csv]

Reports, per failure mode: how many lines were allocated automatically, how many of those were
correct, and how many were sent for review. The number that matters is *wrong* — a receipt
credited to the wrong family is worse than one held back for a human.
"""
import csv, sys
from collections import defaultdict
from openpyxl import load_workbook

AUTO_TIERS = {"M", "A", "B"}
LEDGER = sys.argv[1] if len(sys.argv) > 1 else "examples/fee_ledger.xlsx"
TRUTH = sys.argv[2] if len(sys.argv) > 2 else "examples/ground_truth.csv"

wb = load_workbook(LEDGER, data_only=True)
fam_name = {r[0]: str(r[1] or "") for r in wb["Families"].iter_rows(min_row=4, values_only=True) if r[0]}

got = []
for r in wb["Receipts"].iter_rows(min_row=4, values_only=True):
    if r[0] is None: break
    tier, fid = r[8] or "", r[9] or ""
    # column H is a formula; openpyxl writes no cached value, so derive allocation from the tier
    got.append(dict(date=r[0], amount=float(r[1] or 0), payer=str(r[2] or ""), tier=tier,
                    fid=fid, allocated=tier in AUTO_TIERS and bool(fid)))
truth = list(csv.DictReader(open(TRUTH)))
assert len(truth) == len(got), f"ledger has {len(got)} receipts, truth has {len(truth)} — regenerate"

# the ledger sorts receipts by date, the truth file is in generation order: join on the line itself
def key_g(g): return (str(g["date"])[:10], round(g["amount"], 2), g["payer"].strip().upper()[:23])
def key_t(t): return (str(t["date"])[:10], round(float(t["amount"]), 2), t["payer"].strip().upper()[:23])
by_key = defaultdict(list)
for t in truth: by_key[key_t(t)].append(t)
pairs = []
for g in got:
    bucket = by_key.get(key_g(g))
    assert bucket, f"no ground-truth line for {key_g(g)}"
    pairs.append((g, bucket.pop(0)))

def matches(fid, true_family):
    """the ledger stores a family key; compare on the surname it was built from"""
    if not fid: return False
    have = fam_name.get(fid, "").upper().replace(" / ", "-")
    want = true_family.upper()
    return want in have or have in want

stats = defaultdict(lambda: dict(n=0, auto=0, right=0, wrong=0, review=0))
wrong_rows = []
for g, t in pairs:
    mode = t["failure_mode"].split("+")[0]
    s = stats[mode]; s["n"] += 1
    if not t["true_family"]:
        s["review" if not g["allocated"] else "auto"] += 1; continue
    if g["allocated"]:
        s["auto"] += 1
        if matches(g["fid"], t["true_family"]): s["right"] += 1
        else:
            s["wrong"] += 1
            wrong_rows.append((t["failure_mode"], t["payer"], t["reference"], t["true_family"], g["fid"], fam_name.get(g["fid"], "")))
    else:
        s["review"] += 1

tot = dict(n=0, auto=0, right=0, wrong=0, review=0)
print(f"{'failure mode':18s} {'lines':>6} {'auto':>6} {'correct':>8} {'WRONG':>6} {'review':>7}")
for mode in sorted(stats, key=lambda m: -stats[m]["n"]):
    s = stats[mode]
    for k in tot: tot[k] += s[k]
    print(f"{mode:18s} {s['n']:6d} {s['auto']:6d} {s['right']:8d} {s['wrong']:6d} {s['review']:7d}")
print(f"{'TOTAL':18s} {tot['n']:6d} {tot['auto']:6d} {tot['right']:8d} {tot['wrong']:6d} {tot['review']:7d}")
print()
print(f"allocated automatically : {tot['auto']}/{tot['n']} = {tot['auto']/tot['n']:.0%}")
print(f"of those, correct       : {tot['right']}/{tot['auto']} = {tot['right']/max(tot['auto'],1):.1%}")
print(f"sent for human review   : {tot['review']}")
if wrong_rows:
    print("\nMISALLOCATED:")
    for w in wrong_rows: print("  ", w)
