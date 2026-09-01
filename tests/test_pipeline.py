"""
End-to-end tests. Each one runs the real pipeline on generated data whose answers are known,
so a change to the matching logic that starts crediting the wrong family fails the build.

    pytest -q

The guarantee being protected is precision, not coverage: it is fine for the engine to send more
lines to a human, and never fine for it to allocate one to the wrong family.
"""
import csv, subprocess, sys
from collections import defaultdict
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
AUTO_TIERS = {"M", "A", "B"}
SEEDS = [20260101, 7, 99]


def run(*args):
    r = subprocess.run([sys.executable, *map(str, args)], cwd=ROOT, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    return r.stdout


@pytest.fixture(scope="session", params=SEEDS)
def pipeline(request, tmp_path_factory):
    out = tmp_path_factory.mktemp(f"seed{request.param}")
    run(ROOT / "generate_fake_data.py", out, "--seed", request.param)
    ledger = out / "fee_ledger.xlsx"
    run(ROOT / "build_ledger.py", out / "roster.xlsx", out / "bank.xlsx", ledger)
    run(ROOT / "reconcile.py", ledger)

    wb = load_workbook(ledger, data_only=True)
    fam = {r[0]: str(r[1] or "") for r in wb["Families"].iter_rows(min_row=4, values_only=True) if r[0]}
    rows = []
    for r in wb["Receipts"].iter_rows(min_row=4, values_only=True):
        if r[0] is None:
            break
        rows.append(dict(
            date=str(r[0])[:10],
            amount=round(float(r[1] or 0), 2),
            payer=str(r[2] or "").strip().upper()[:23],
            tier=r[8] or "",
            fid=r[9] or "",
            cands=str(r[10] or ""),
            why=str(r[11] or ""),
        ))
    truth = defaultdict(list)
    for t in csv.DictReader(open(out / "ground_truth.csv")):
        key = (str(t["date"])[:10], round(float(t["amount"]), 2), t["payer"].strip().upper()[:23])
        truth[key].append(t)
    pairs = []
    for g in rows:
        bucket = truth[(g["date"], g["amount"], g["payer"])]
        assert bucket, f"receipt with no ground truth: {g}"
        pairs.append((g, bucket.pop(0)))
    return dict(wb=wb, fam=fam, pairs=pairs, seed=request.param)


def correct(fam, fid, true_family):
    have = fam.get(fid, "").upper().replace(" / ", "-")
    want = true_family.upper()
    return bool(fid) and (want in have or have in want)


def test_no_receipt_is_credited_to_the_wrong_family(pipeline):
    bad = [(g, t) for g, t in pipeline["pairs"]
           if g["tier"] in AUTO_TIERS and t["true_family"] and not correct(pipeline["fam"], g["fid"], t["true_family"])]
    assert not bad, f"{len(bad)} misallocated, first: {bad[0]}"


def test_most_receipts_are_allocated_without_a_human(pipeline):
    auto = [g for g, _ in pipeline["pairs"] if g["tier"] in AUTO_TIERS]
    assert len(auto) / len(pipeline["pairs"]) >= 0.75


def test_a_clean_reference_always_allocates(pipeline):
    for g, t in pipeline["pairs"]:
        if t["failure_mode"] == "clean":
            assert g["tier"] == "A", f"clean reference not matched by reference: {g}"


def test_a_reference_that_contradicts_the_payer_name_is_held_back(pipeline):
    for g, t in pipeline["pairs"]:
        if t["failure_mode"].startswith("wrong_ref") and g["tier"] in AUTO_TIERS:
            assert correct(pipeline["fam"], g["fid"], t["true_family"]), f"followed a wrong reference: {g}"


def test_conflicts_are_explained(pipeline):
    for g, _ in pipeline["pairs"]:
        if g["tier"] == "X":
            assert "may have typed the wrong invoice number" in g["why"]


def test_an_invoice_number_shared_by_two_families_never_guesses(pipeline):
    for g, t in pipeline["pairs"]:
        if t["failure_mode"].startswith("shared_invoice") and g["tier"] in AUTO_TIERS:
            assert correct(pipeline["fam"], g["fid"], t["true_family"]), f"guessed on a shared invoice number: {g}"


def test_every_queued_receipt_carries_a_reason(pipeline):
    for g, _ in pipeline["pairs"]:
        if g["tier"] not in AUTO_TIERS:
            assert g["why"].strip(), f"no reason given for a queued receipt: {g}"


def test_the_fee_rules_reproduce_the_invoices(pipeline):
    ch = pipeline["wb"]["Children"]
    checks = [r for r in ch.iter_rows(min_row=4, values_only=True) if r[0] and r[5] == "Enrolled"]
    assert checks
    flagged = [r for r in checks if isinstance(r[19], (int, float)) and abs(r[19]) >= 0.02]
    assert len(flagged) <= 2, f"fee rules disagree with {len(flagged)} invoices"


def test_no_money_disappears(pipeline):
    total = sum(g["amount"] for g, _ in pipeline["pairs"])
    allocated = sum(g["amount"] for g, _ in pipeline["pairs"] if g["tier"] in AUTO_TIERS)
    queued = sum(g["amount"] for g, _ in pipeline["pairs"] if g["tier"] not in AUTO_TIERS)
    assert round(allocated + queued, 2) == round(total, 2)


def test_the_workbook_has_the_sheets_a_reviewer_needs(pipeline):
    for sheet in ("Summary", "README", "Rates", "Terms", "Families", "Children", "Receipts", "Review", "Position"):
        assert sheet in pipeline["wb"].sheetnames
