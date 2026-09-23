"""
preflight.py - check a real roster and bank export against what build_ledger.py requires.

    python3 preflight.py STUDENTS_2026.xlsx BANK.xlsx

build_ledger.py reads the roster by heading, under any alias listed in sources.py, on whichever
sheet and row of the workbook carries those headings. It reads the bank export by heading when the
file has one, else by column position. A roster whose headings are spelled some new way is skipped
entirely; a headerless bank export with the columns in a different order produces wrong answers
silently, which is worse.

This reports what it found, what is missing, and what to do about it. It changes nothing.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from terms import TERM_IDS, default_term   # noqa: E402  - the one place a term is defined
from sources import (ROSTER_COLS, ROSTER_REQUIRED, HEADER_SCAN_ROWS, BANK_COLS, BANK_POSITIONAL,   # noqa: E402
                     locate_roster, bank_header)

# The term a roster is checked against: the same one build_ledger.py would build by default.
DEFAULT_TERM = default_term()
YEAR = DEFAULT_TERM[-4:]

# The roster is read by heading through sources.ROSTER_COLS; any listed alias will do.
ROSTER_WHAT = {
    "status":   'enrolment status. "Inscrit" is enrolled; blank with an invoice and a fee is a late enrolment.',
    "pupil":    "pupil name. SURNAME in capitals first, then first name; first-name-first rows are read the other way round.",
    "invoice":  "invoice number for the term, e.g. 2026-004.",
    "fees":     "tuition invoiced. Blank means the child was not invoiced.",
    "supplies": "supplies charge. A negative value is read as a discount.",
    "reg":      "one-off registration fee, or blank.",
    "paid":     'free text, e.g. "paid" / "part paid".',
}

# The bank file is read by heading when its first row has one, else positionally as sources.BANK_POSITIONAL.
BANK_POSITIONS = BANK_POSITIONAL

OK, BAD, WARN = "  ok  ", " FAIL ", " warn "


def line(status, text):
    print(f"[{status}] {text}")


def check_roster(path):
    print(f"\n=== roster: {path} ===")
    try:
        where = locate_roster(path)
    except Exception as e:
        line(BAD, f"could not read as .xlsx: {e}")
        return False
    if where is None:
        need = ", ".join(ROSTER_COLS[f][0] for f in ROSTER_REQUIRED)
        line(BAD, f"no sheet has a heading row naming {need} in its first {HEADER_SCAN_ROWS} rows")
        print(f"        sheets in this workbook: {pd.ExcelFile(path).sheet_names}")
        print("        build_ledger.py takes the first sheet whose top rows carry those headings."
              " Check the roster sheet is in this file and its headings are spelled as above;"
              " a new spelling is one alias added to sources.ROSTER_COLS.")
        return False
    sheet, hdr = where
    df = pd.read_excel(path, sheet_name=sheet, header=hdr)
    df.columns = [str(c).strip() for c in df.columns]
    line(OK, f'roster on sheet "{sheet}", headings on row {hdr + 1}: {len(df)} rows, {len(df.columns)} columns')

    ok = True
    col = {}
    for field_, aliases in ROSTER_COLS.items():
        found = next((a for a in aliases if a in df.columns), None)
        if found:
            col[field_] = found
            line(OK, f'{field_}: column "{found}" found')
        elif field_ in ROSTER_REQUIRED:
            ok = False
            line(BAD, f'{field_}: none of {", ".join(aliases)} present - {ROSTER_WHAT[field_]}')
        else:
            line(WARN, f'{field_}: none of {", ".join(aliases)} present - optional; {ROSTER_WHAT[field_]}')

    extra = [c for c in df.columns if c not in col.values() and not c.startswith("Unnamed")]
    if extra:
        line(WARN, f"columns present but never read: {', '.join(extra[:8])}")
        print("        (harmless - but check none of them is your real invoice/fee column"
              " under a different name)")

    if "status" in col:
        raw = df[col["status"]]
        counts = raw.dropna().astype(str).str.strip().value_counts()
        enrolled, blank = int(counts.get("Inscrit", 0)), int(raw.isna().sum())
        if enrolled:
            line(OK, f'{enrolled} rows have {col["status"]} = "Inscrit" and will be treated as enrolled'
                     + (f"; {blank} with a blank status are enrolled if they carry an invoice and a fee" if blank else ""))
        else:
            ok = False
            line(BAD, f'no row has {col["status"]} = "Inscrit" - every pupil would be dropped')
            print(f"        values actually present: {dict(list(counts.items())[:6])}")

    if "pupil" in col:
        sample = [str(v) for v in df[col["pupil"]].dropna().head(3)]
        bad = [s for s in sample if not any(t.isupper() and len(t) > 1 for t in s.split())]
        if bad:
            line(WARN, "pupil names may not parse - the surname must be in CAPITALS, first")
            print(f"        found: {bad}")
        else:
            line(OK, f"pupil name format looks parseable, e.g. {sample[:2]}")

    if "invoice" in col:
        refs = df[col["invoice"]].dropna().astype(str).str.strip()
        matching = refs.str.match(rf"^\s*{YEAR}\s*-?\s*\d{{3}}").sum()
        if len(refs) and matching == 0:
            ok = False
            line(BAD, f"no invoice number matches the {YEAR}-nnn pattern the engine looks for")
            print(f"        found instead: {list(refs.head(3))}")
            print(f"        both engine.py and build_ledger.py build their reference patterns"
                  f" from the term, and this check assumed {DEFAULT_TERM}. If the roster is for"
                  f" another term, add it to terms.py and pass --term.")
        elif len(refs):
            line(OK, f"{matching}/{len(refs)} invoice numbers match the {YEAR}-nnn pattern")
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

    if len(raw) and bank_header(raw.iloc[0]):
        # A headed export (Barclays) is read by name, so column order does not matter.
        head = [str(x).strip() for x in raw.iloc[0]]
        ok = True
        for field_, aliases in BANK_COLS.items():
            found = next((a for a in aliases if a in head), None)
            if found:
                line(OK, f'{field_}: column "{found}" found (read by heading)')
            elif field_ != "cat":
                ok = False
                line(BAD, f'{field_}: no column named {" or ".join(aliases)} in the header row')
        return ok

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
    print(f"\n        Terms are defined in one place, the TERMS list in terms.py, which currently"
          f"\n        holds: {', '.join(TERM_IDS)}. This check assumed {DEFAULT_TERM}."
          f"\n        For another term, add an entry there and pass build_ledger.py --term <ID>.")
    return 0 if (r_ok and b_ok) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
