"""
namecheck.py - before a push: is any name from the real files in a tracked file or a commit message?

    python3 namecheck.py ROSTER.xlsx BANK.xls [--base origin/main] [--ignore WORD,WORD]

Takes every surname on the roster and every word of every payer name on a receipt in the bank
export, and looks for them in the tracked text files and in the messages of the commits since
--base. Prints one line per name found, with where, and exits 1 if anything was.

It exists because the harness checks allocations and .gitignore blocks spreadsheets, and neither
stopped a real surname reaching a Markdown file and a commit message. Invented names are fine; an
invented name that happens to be a real one on this term's roster is not, and this finds those too.

Some payer words are plain words (a parent called GREEN, a colour in the code). --ignore takes
those out, and they are printed in the summary so whoever reviews the push sees what was waved
through. Nothing from the real files is printed except a token that is already in a tracked file.
"""
import re
import subprocess
import sys
import unicodedata
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
from sources import ROSTER_COLS, BANK_COLS, NAME_NOISE, locate_roster, bank_header   # noqa: E402

TEXT = {".py", ".md", ".yml", ".yaml", ".sh", ".txt", ".toml", ".cfg", ".ini"}
# Words that sit inside payer names and roster cells without being anyone's name.
STOP = {"FEES", "FEE", "SCHOOL", "FRENCH", "CLUB", "LTD", "LIMITED", "BANK", "PAYMENT", "TRANSFER",
        "TERM", "MRS", "MISS", "AND", "THE", "FOR", "FROM", "TWINS", "FAMILY", "PARENTS", "TOTAL",
        "INSCRIT", "ACCOUNT", "TRUST", "SERVICES", "HOLDINGS", "CLASSE"} | set(NAME_NOISE)
MIN_LEN = 4
SHOW = 6   # locations listed per name before "(+n more)"


def norm(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper()


def surnames(roster):
    """Surname tokens on the roster: the capitalised leading tokens of each pupil cell, else the last."""
    where = locate_roster(roster)
    if where is None:
        sys.exit(f"{roster}: no roster sheet found")
    df = pd.read_excel(roster, sheet_name=where[0], header=where[1])
    df.columns = [str(c).strip() for c in df.columns]
    col = next(a for a in ROSTER_COLS["pupil"] if a in df.columns)
    out = set()
    for cell in df[col].dropna().astype(str):
        toks = re.sub(r"\s*-\s*", "-", cell).split("(")[0].split()
        caps = [t for t in toks if t.isupper() and len(t) > 1]
        for t in (caps or toks[-1:]):
            out.update(re.split(r"[^A-Za-zÀ-ÿ']+", t))
    return {norm(t) for t in out if len(t) >= MIN_LEN}


def payers(bank):
    """Every word of every payer name on a receipt: the first 23 characters of the memo, the bank's
    payer slot, on lines with a positive amount. Outgoing lines are the school's own spending and
    never reach the ledger, so their words (BILL, INTEREST, a supplier) are not names to protect."""
    raw = pd.read_excel(bank, header=None)
    if len(raw) and bank_header(raw.iloc[0]):
        head = [str(x).strip() for x in raw.iloc[0]]
        memo = raw.iloc[1:, head.index(next(a for a in BANK_COLS["memo"] if a in head))]
        amount = raw.iloc[1:, head.index(next(a for a in BANK_COLS["amount"] if a in head))]
    else:
        memo, amount = raw.iloc[:, 2], raw.iloc[:, 1]
    amount = pd.to_numeric(amount.astype(str).str.replace(r"[£,\s]", "", regex=True), errors="coerce")
    out = set()
    for m in memo[amount > 0].dropna().astype(str):
        out.update(re.split(r"[^A-Za-z']+", m[:23]))
    return {norm(t) for t in out if len(t) >= MIN_LEN}


def pattern(tokens):
    """One regex for all tokens, whole-word, longest first so a hyphenated surname beats its parts."""
    alts = "|".join(re.escape(t) for t in sorted(tokens, key=len, reverse=True))
    return re.compile(rf"(?<![A-Z])(?:{alts})(?![A-Z])")


def in_files(rx, found):
    files = subprocess.run(["git", "ls-files"], cwd=ROOT, capture_output=True, text=True).stdout.split()
    for f in files:
        if Path(f).suffix not in TEXT or f.startswith("examples/"):
            continue
        text = (ROOT / f).read_text(encoding="utf-8", errors="replace")
        for n, line in enumerate(text.splitlines(), 1):
            for tok in set(rx.findall(norm(line))):
                found.setdefault(tok, []).append(f"{f}:{n}")


def in_messages(rx, base, found):
    ok = subprocess.run(["git", "rev-parse", "--verify", "-q", base], cwd=ROOT, capture_output=True)
    if ok.returncode != 0:
        print(f"(no ref {base}: commit messages not checked)")
        return
    log = subprocess.run(["git", "log", "--format=%h%x00%B%x01", f"{base}..HEAD"],
                         cwd=ROOT, capture_output=True, text=True).stdout
    for entry in log.split("\x01"):
        if "\x00" not in entry:
            continue
        sha, body = entry.split("\x00", 1)
        for tok in set(rx.findall(norm(body))):
            found.setdefault(tok, []).append(f"commit {sha.strip()}")


def parse(argv):
    paths, base, ignore = [], "origin/main", set()
    i = 1
    while i < len(argv):
        if argv[i] == "--base" and i + 1 < len(argv):
            base, i = argv[i + 1], i + 2
        elif argv[i] == "--ignore" and i + 1 < len(argv):
            ignore, i = {norm(w) for w in argv[i + 1].split(",") if w.strip()}, i + 2
        else:
            paths.append(argv[i])
            i += 1
    return paths, base, ignore


def main(argv):
    paths, base, ignore = parse(argv)
    if len(paths) < 2:
        print(__doc__)
        return 2
    tokens = (surnames(paths[0]) | payers(paths[1])) - STOP - ignore
    found = {}
    if tokens:
        rx = pattern(tokens)
        in_files(rx, found)
        in_messages(rx, base, found)
    for tok in sorted(found):
        where = found[tok]
        more = f" (+{len(where) - SHOW} more)" if len(where) > SHOW else ""
        print(f"{tok:<20} {', '.join(where[:SHOW])}{more}")
    waved = f"; ignored on request: {', '.join(sorted(ignore))}" if ignore else ""
    print(f"{len(tokens)} names checked against tracked files and commits since {base}: "
          f"{len(found)} found{waved}")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
