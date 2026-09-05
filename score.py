"""
score.py - measure the engine against the generator's ground truth.

    python score.py [ledger.xlsx] [ground_truth.csv]

Reports, per failure mode: how many lines were allocated automatically, how many of those were
correct, and how many were sent for review. The number that matters is wrong - a receipt
credited to the wrong family is worse than one held back for a human.
"""
import csv, re, sys, unicodedata
from collections import Counter, defaultdict
from openpyxl import load_workbook

from engine import Config, load_roster, norm, tokens

AUTO_TIERS = {"M", "A", "B"}
LEDGER = sys.argv[1] if len(sys.argv) > 1 else "examples/fee_ledger.xlsx"
TRUTH = sys.argv[2] if len(sys.argv) > 2 else "examples/ground_truth.csv"

wb = load_workbook(LEDGER, data_only=True)
roster = load_roster(wb, Config())


def norm_family(s):
    """Family name to a comparable form: accents stripped, ' / ' joined with '-', upper case.

    Comparison is exact on this form. Substring comparison is wrong here: MARCHETTI is a
    substring of ACHTERBERG-MARCHETTI, and those are two unrelated families.
    """
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.upper().replace(" / ", "-").replace("/", "-")
    return re.sub(r"\s+", " ", s).strip()


def invoices_of(cell):
    return {p.strip() for p in re.split(r"[;,]", str(cell or "")) if p.strip()}


fam_name, fam_invoices = {}, {}
for r in wb["Families"].iter_rows(min_row=4, values_only=True):
    if not r[0]:
        continue
    fam_name[r[0]] = norm_family(r[1])
    fam_invoices[r[0]] = invoices_of(r[4])

# Families that share a name cannot be told apart by name alone; the invoice number breaks the tie.
name_counts = Counter(fam_name.values())

got = []
for r in wb["Receipts"].iter_rows(min_row=4, values_only=True):
    if r[0] is None:
        break
    tier, fid = r[8] or "", r[9] or ""
    got.append(dict(
        date=r[0], amount=float(r[1] or 0), payer=str(r[2] or ""), memo=str(r[14] or ""),
        tier=tier, fid=fid, allocated=tier in AUTO_TIERS and bool(fid),
    ))

truth = list(csv.DictReader(open(TRUTH)))
assert len(truth) == len(got), f"ledger has {len(got)} receipts, truth has {len(truth)} - regenerate"


def key_g(g):
    return (str(g["date"])[:10], round(g["amount"], 2), g["payer"].strip().upper()[:23])


def key_t(t):
    return (str(t["date"])[:10], round(float(t["amount"]), 2), t["payer"].strip().upper()[:23])


by_key = defaultdict(list)
for t in truth:
    by_key[key_t(t)].append(t)
pairs = []
for g in got:
    bucket = by_key.get(key_g(g))
    assert bucket, f"no ground-truth line for {key_g(g)}"
    pairs.append((g, bucket.pop(0)))


def matches(fid, true_family, true_invoice=""):
    """True when fid is the family the generator says the receipt belongs to.

    Exact on the normalised name. Where two families share a name, the ground-truth invoice
    number must also be one of that family's invoices.
    """
    if not fid:
        return False
    have = fam_name.get(fid, "")
    want = norm_family(true_family)
    if not have or have != want:
        return False
    if name_counts[have] > 1 and true_invoice:
        return true_invoice.strip() in fam_invoices.get(fid, set())
    return True


stats = defaultdict(lambda: dict(n=0, auto=0, right=0, wrong=0, review=0))
wrong_rows = []
for g, t in pairs:
    mode = t["failure_mode"].split("+")[0]
    s = stats[mode]
    s["n"] += 1
    if not t["true_family"]:
        s["review" if not g["allocated"] else "auto"] += 1
        continue
    if g["allocated"]:
        s["auto"] += 1
        if matches(g["fid"], t["true_family"], t.get("true_invoice", "")):
            s["right"] += 1
        else:
            s["wrong"] += 1
            wrong_rows.append((t["failure_mode"], t["payer"], t["reference"], t["true_family"], g["fid"], fam_name.get(g["fid"], "")))
    else:
        s["review"] += 1

tot = dict(n=0, auto=0, right=0, wrong=0, review=0)
print(f"{'failure mode':18s} {'lines':>6} {'auto':>6} {'correct':>8} {'WRONG':>6} {'review':>7}")
for mode in sorted(stats, key=lambda m: -stats[m]["n"]):
    s = stats[mode]
    for k in tot:
        tot[k] += s[k]
    print(f"{mode:18s} {s['n']:6d} {s['auto']:6d} {s['right']:8d} {s['wrong']:6d} {s['review']:7d}")
print(f"{'TOTAL':18s} {tot['n']:6d} {tot['auto']:6d} {tot['right']:8d} {tot['wrong']:6d} {tot['review']:7d}")
print()
print(f"allocated automatically : {tot['auto']}/{tot['n']} = {tot['auto']/tot['n']:.0%}")
print(f"of those, correct       : {tot['right']}/{tot['auto']} = {tot['right']/max(tot['auto'],1):.1%}")
print(f"sent for human review   : {tot['review']}")

# ---------------------------------------------------------------------------- baseline
# A number with nothing to compare it against says very little. This is the simplest thing that
# could work: allocate when exactly one roster surname appears in the memo, and never otherwise.
# No references, no aliasing, no tie-breaks. Whatever the engine scores above this is what the
# tier ladder is actually buying.
base_auto = base_right = 0
for g, t in pairs:
    hits = {f for tok in set(tokens(norm(g["memo"]))) for f in roster.sur_index.get(tok, set())}
    if len(hits) != 1:
        continue
    base_auto += 1
    if t["true_family"] and matches(next(iter(hits)), t["true_family"], t.get("true_invoice", "")):
        base_right += 1

print()
print(f"{'':24s} {'allocated':>12s} {'correct':>10s}")
print(f"{'naive surname-only':24s} {base_auto:>5d}/{tot['n']:<6d} {base_right/max(base_auto,1):>9.1%}")
print(f"{'engine':24s} {tot['auto']:>5d}/{tot['n']:<6d} {tot['right']/max(tot['auto'],1):>9.1%}")
print(f"{'the ladder is worth':24s} {tot['auto']-base_auto:>+5d} receipt(s)")
if wrong_rows:
    print("\nMISALLOCATED:")
    for w in wrong_rows:
        print("  ", w)
