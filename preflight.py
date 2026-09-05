"""
preflight.py - check a real roster and bank export against what build_ledger.py requires.

    python3 preflight.py STUDENTS_2026.xlsx BANK.xlsx

build_ledger.py reads both files by exact column name (roster) and by column position (bank).
Neither adapts. A roster with the right data under the wrong heading raises KeyError; a bank
export with the columns in a different order produces wrong answers silently, which is worse.

This reports what it found, what is missing, and what to do about it. It changes nothing.
"""
import sys
from pathlib import Path

import pandas as pd

# build_ledger.py:29-57. The roster is read by name; every one of these must be present.
ROSTER_REQUIRED = {
    "Statut:": 'enrolment status. Rows counted as enrolled are exactly those equal to "Inscrit".',
    "LES ELEVES": "pupil name. SURNAME in capitals first, then first name.",
    "INVOICE NUMBER": "invoice number for the term, e.g. 2026-004.",
    "FEES": "tuition invoiced. Blank means the child was not invoiced.",
    "OFFICE SUPPLIES": "supplies charge. A negative value is read as a discount.",
    "REG FEES ONE OFF": "one-off registration fee, or blank.",
    "PAID": 'free text, e.g. "paid" / "part paid".',
}

# build_ledger.py:125-128. The bank file is read with header=None and named positionally.
BANK_POSITIONS = ["date", "amount", "memo", "cat", "note", "extra"]

OK, BAD, WARN = "  ok  ", " FAIL ", " warn "


def line(status, text):
    print(f"[{status}] {text}")


def check_roster(path):
    print(f"\n=== roster: {path} ===")
    try:
        df = pd.read_excel(path)
    except Exception as e:
        line(BAD, f"could not read as .xlsx: {e}")
        return False
    df.columns = [str(c).strip() for c in df.columns]
    line(OK, f"read {len(df)} rows, {len(df.columns)} columns")

    ok = True
    for col, what in ROSTER_REQUIRED.items():
        if col in df.columns:
            line(OK, f'column "{col}" found')
        else:
            ok = False
            line(BAD, f'column "{col}" MISSING - {what}')

    extra = [c for c in df.columns if c not in ROSTER_REQUIRED and not c.startswith("Unnamed")]
    if extra:
        line(WARN, f"columns present but never read: {', '.join(extra[:8])}")
        print("        (harmless - but check none of them is your real invoice/fee column"
              " under a different name)")

    if "Statut:" in df.columns:
        counts = df["Statut:"].astype(str).str.strip().value_counts()
        enrolled = int(counts.get("Inscrit", 0))
        if enrolled:
            line(OK, f'{enrolled} rows have Statut: = "Inscrit" and will be treated as enrolled')
        else:
            ok = False
            line(BAD, 'no row has Statut: = "Inscrit" - every pupil would be dropped')
            print(f"        values actually present: {dict(list(counts.items())[:6])}")

    if "LES ELEVES" in df.columns:
        sample = [str(v) for v in df["LES ELEVES"].dropna().head(3)]
        bad = [s for s in sample if not any(t.isupper() and len(t) > 1 for t in s.split())]
        if bad:
            line(WARN, "pupil names may not parse - the surname must be in CAPITALS, first")
            print(f"        found: {bad}")
        else:
            line(OK, f"pupil name format looks parseable, e.g. {sample[:2]}")

    if "INVOICE NUMBER" in df.columns:
        refs = df["INVOICE NUMBER"].dropna().astype(str).str.strip()
        matching = refs.str.match(r"^\s*2026\s*-?\s*\d{3}").sum()
        if len(refs) and matching == 0:
            ok = False
            line(BAD, "no invoice number matches the 2026-nnn pattern the engine looks for")
            print(f"        found instead: {list(refs.head(3))}")
            print("        engine.py builds its reference patterns from Config.term_cur;"
                  " build_ledger.py:22 hard-codes JAN-2026 / AUT-2025.")
        elif len(refs):
            line(OK, f"{matching}/{len(refs)} invoice numbers match the 2026-nnn pattern")
    return ok


def check_bank(path):
    print(f"\n=== bank export: {path} ===")
    try:
        raw = pd.read_excel(path, header=None)
    except Exception as e:
        line(BAD, f"could not read as .xlsx: {e}")
        return False
    line(OK, f"read {len(raw)} rows, {len(raw.columns)} columns")

    if len(raw.columns) < 3:
        line(BAD, f"need at least 3 columns (date, amount, memo), found {len(raw.columns)}")
        return False

    print("\n        read positionally as:", ", ".join(BANK_POSITIONS[:len(raw.columns)]))
    print("        first row of your file:")
    for i, v in enumerate(raw.iloc[0][:6]):
        name = BANK_POSITIONS[i] if i < len(BANK_POSITIONS) else f"col{i}"
        print(f"          {name:8s} = {str(v)[:56]}")

    ok = True
    col0 = pd.to_datetime(raw[0], errors="coerce")
    if col0.notna().mean() > 0.8:
        line(OK, "column 1 parses as dates")
    else:
        ok = False
        line(BAD, "column 1 does not parse as dates - the columns are in the wrong order")

    col1 = pd.to_numeric(raw[1], errors="coerce")
    if col1.notna().mean() > 0.8:
        line(OK, "column 2 parses as numbers")
    else:
        ok = False
        line(BAD, "column 2 does not parse as amounts - the columns are in the wrong order")

    memos = raw[2].dropna().astype(str)
    if len(memos) and memos.str.len().mean() > 5:
        line(OK, "column 3 looks like payer/reference text")
        print(f"        e.g. {list(memos.head(2))}")
    else:
        ok = False
        line(BAD, "column 3 does not look like a payment memo")

    if str(raw.iloc[0][2]).strip().lower() in {"memo", "reference", "description", "details"}:
        ok = False
        line(BAD, "row 1 looks like a HEADER ROW - delete it. The file is read with header=None,"
                  " so a header row is treated as a payment.")
    return ok


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        return 2
    roster, bank = Path(argv[1]), Path(argv[2])
    for p in (roster, bank):
        if not p.exists():
            line(BAD, f"file not found: {p}")
            return 2

    r_ok = check_roster(roster)
    b_ok = check_bank(bank)

    print("\n=== verdict ===")
    if r_ok and b_ok:
        line(OK, "both files match what build_ledger.py expects. Next:")
        print(f"        python3 build_ledger.py {roster} {bank} fee_ledger.xlsx")
        print(f"        python3 reconcile.py fee_ledger.xlsx")
        print("\n        score.py will NOT work - it needs a ground-truth answer key,"
              " which real data does not have.")
    else:
        line(BAD, "not ready. Fix the FAIL lines above - rename the roster columns to match,"
                  " and reorder the bank columns to date / amount / memo.")
    print("\n        Whatever you do, note the term is hard-coded to JAN-2026 in six places in"
          "\n        build_ledger.py (line 22, and the patterns and dates at 150-159, 190, 289-291).")
    return 0 if (r_ok and b_ok) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
