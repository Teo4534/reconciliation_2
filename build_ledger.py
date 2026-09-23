"""
build_ledger.py - turn a pupil roster and a bank export into a structured fee ledger.

    python build_ledger.py <roster.xlsx> <bank.xlsx> <ledger_out.xlsx> [--term JAN-2026]

Writes Rates, Terms, Families, Children and Receipts. The fee rules live in Rates as data,
so a price change is a cell edit, not a code change. Run reconcile.py next to allocate receipts.

Every term-specific fact lives in terms.py. Running a different term is a --term argument, not an
edit: nothing in this file names a year.
"""
import re, sys, unicodedata
from collections import defaultdict
from difflib import SequenceMatcher
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

from terms import default_term, resolve
from engine import Config as EngineConfig
from sources import (ROSTER_COLS, ROSTER_REQUIRED, HEADER_SCAN_ROWS, BANK_COLS, BANK_POSITIONAL,
                     NOT_A_RECEIPT, INVOICE_BLANKS, NAME_NOISE, locate_roster, bank_header)

def parse_args(argv):
    """Split '<roster> <bank> <out> [--term ID]' into three paths and the chosen term entry.

    The term defaults to terms.default_term(), so existing commands keep working. The entry before
    the chosen one is the prior term, which is what receipts are compared against.
    """
    paths, want = [], default_term()
    i = 0
    while i < len(argv):
        if argv[i] == "--term":
            if i + 1 >= len(argv):
                sys.exit("--term needs a term ID, e.g. --term JAN-2026")
            want, i = argv[i + 1], i + 2
            continue
        paths.append(argv[i])
        i += 1
    while len(paths) < 3:
        paths.append(["examples/roster.xlsx", "examples/bank.xlsx", "examples/fee_ledger.xlsx"][len(paths)])
    try:
        prior, cur = resolve(want)
    except ValueError as e:
        sys.exit(str(e))
    return paths[0], paths[1], paths[2], prior, cur


ROSTER, BANK, OUT, PRIOR, CUR = parse_args(sys.argv[1:])
TERM_CUR, TERM_PRIOR = CUR["id"], PRIOR["id"]
YEAR_CUR, YEAR_PRIOR = TERM_CUR[-4:], TERM_PRIOR[-4:]

# ================================================================ source files
# Headings, filters and noise words live in sources.py, shared with preflight.py. A stray letter in
# the invoice year is a typo for the same invoice.
INVOICE_FIXES = ((re.compile(r"^(\d{4})[A-Za-z]+-"), r"\1-"),)
# When a term and the one before it share a year, the invoice series in terms.py is what tells
# their references apart: autumn 2026 issues 2026-5xx where January 2026 issued 2026-0xx. Read
# from the same text the Terms sheet shows, so build_ledger.py and reconcile.py cannot disagree.
SAME_YEAR = YEAR_CUR == YEAR_PRIOR
CUR_SERIES = EngineConfig.series_bounds(CUR["series"]) if SAME_YEAR else ()
PRIOR_SERIES = EngineConfig.series_bounds(PRIOR["series"]) if SAME_YEAR else ()

def in_series(num, lo_hi):
    """Is this three-digit part inside a term's invoice series? True when no series is set."""
    return not lo_hi or lo_hi[0] <= int(num) <= lo_hi[1]

def norm(x):
    return unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode()

# ---------------------------------------------------------------- roster
def pick(df, field_):
    """The actual column heading in `df` for a canonical field, or None if the file has not got it."""
    for name in ROSTER_COLS[field_]:
        if name in df.columns:
            return name
    return None

# The roster may arrive as the office's whole workbook: the right sheet among several, with rows of
# teacher and class names above the headings. Find it rather than assume sheet 1, row 1.
where = locate_roster(ROSTER)
if where is None:
    top = pd.read_excel(ROSTER, header=None, nrows=HEADER_SCAN_ROWS)
    seen = [str(v) for v in top.iloc[0].dropna()][:8] if len(top) else []
    need = ", ".join(ROSTER_COLS[f][0] for f in ROSTER_REQUIRED)
    sys.exit(f"{ROSTER}: no sheet has a heading row naming {need} in its first {HEADER_SCAN_ROWS} rows. "
             f"First sheet starts: {seen}")
ROSTER_SHEET, ROSTER_HEADER = where
s = pd.read_excel(ROSTER, sheet_name=ROSTER_SHEET, header=ROSTER_HEADER)
s.columns = [str(c).strip() for c in s.columns]
C = {f: pick(s, f) for f in ROSTER_COLS}
# A spreadsheet carries its own furniture: a grand-total row at the foot, blank spacer rows, and
# (in the real export) thousands of empty columns Excel invented. Keep only rows that name a pupil.
s = s[s[C["pupil"]].notna() & (s[C["status"]].astype(str).str.strip() != "Total:")].copy()
# Free-text notes sit in unlabelled columns whose position moves between exports. Take the ones
# holding text; a stray unlabelled number is the total row's spill, not a note.
note_cols = [c for c in s.columns if str(c).startswith("Unnamed")
             and s[c].map(lambda v: isinstance(v, str) and v.strip() != "").any()]
# Enrolment. "Inscrit" is explicit, but the real roster leaves the status blank on rows added after
# the list was first drawn up - a sibling tacked onto an existing invoice, and the block of late
# joiners at the foot. A blank status with an invoice number and a fee is an enrolled child: the
# office invoiced them. Dropping those loses their payments, which is the failure this guards.
status_raw = s[C["status"]].astype(str).str.strip() if C["status"] else pd.Series("", index=s.index)
status_blank = (s[C["status"]].isna() | (status_raw == "")) if C["status"] else pd.Series(True, index=s.index)
has_invoice = s[C["invoice"]].notna() if C["invoice"] else False
has_fee = s[C["fees"]].notna() if C["fees"] else False
s["_late_join"] = status_blank & has_invoice & has_fee
s = s[(status_raw == "Inscrit") | s["_late_join"]].copy().reset_index(drop=True)

def parse_name(raw):
    """Split a roster cell into (SURNAME, first name).

    The roster's own convention is the surname first in capitals. Rows added later are typed in
    ordinary case as "Chloe Renard", first name first, and reading those the house way would make
    CHLOE the surname and lose every payment that spells the family's name. So when no token is
    capitalised, the last token is taken as the surname (with a lower-case particle before it).
    """
    t = re.sub(r"\s*-\s*", "-", str(raw)).strip()
    t = t.split("(")[0]
    toks = t.split()
    sur = []
    for tok in toks:
        if tok.isupper() and len(tok) > 1: sur.append(tok)
        else: break
    if not sur:
        clean = [tok for tok in toks if re.search(r"[A-Za-zÀ-ÿ]", tok) and tok.upper() not in NAME_NOISE]
        if len(clean) >= 2:
            # Title case throughout, so the roster's surname-first rule does not apply. The last
            # token is the surname. A middle name must not join it: "Grace Bayo Chen" is CHEN, or
            # BAYO becomes surname evidence for a child who is not a Bayo. A lower-case tail is a
            # particle, so the token before it is pulled in: "Aisha Ndiaye diop" is NDIAYE DIOP.
            cut = len(clean) - 1
            while cut > 1 and clean[cut][:1].islower():
                cut -= 1
            sur, first = clean[cut:], clean[0]
        else:
            sur, first = ([toks[0]], toks[1] if len(toks) > 1 else "")
    else:
        rest = toks[len(sur):]
        first = rest[0] if rest else ""
    first = re.sub(r"[^A-Za-zÀ-ÿ'\-]", "", first)
    sur = [t for t in sur if t.upper() not in NAME_NOISE] or sur
    return " ".join(sur).upper(), first

rows = []
for i, r in s.iterrows():
    cell = lambda f: r[C[f]] if C[f] else None
    sur, first = parse_name(cell("pupil"))
    # Excel forces a text invoice number with a leading apostrophe ('2026-999); strip it, and the
    # 'a' suffix the roster puts on a sibling's row, so siblings share one invoice key.
    inv_raw = "" if pd.isna(cell("invoice")) else str(cell("invoice")).strip().lstrip("'")
    if inv_raw.lower() in INVOICE_BLANKS:
        inv_raw = ""
    for pat, sub_ in INVOICE_FIXES:
        inv_raw = pat.sub(sub_, inv_raw)
    inv = re.sub(r"[a.]$", "", inv_raw)
    note = " | ".join(str(r[c]).strip() for c in note_cols if pd.notna(r[c]))
    if r["_late_join"]:
        note = " | ".join(x for x in [note, "status blank on roster  -  taken as a late enrolment"] if x)
    # The roster's convention is SURNAME first in capitals. Where a row does not follow it the
    # surname is inferred, so say so: an inferred surname is what the bank memo is matched against.
    inferred = not re.match(r"^\S*[A-ZÀ-Ý]{2,}\S*(\s|$)", str(cell("pupil")).strip())
    if inferred:
        note = " | ".join(x for x in [note, f"name not in SURNAME-first form  -  read as surname '{sur}'; confirm"] if x)
    paid = "" if pd.isna(cell("paid")) else str(cell("paid")).strip()
    fees = cell("fees"); sup = cell("supplies"); reg = cell("reg")
    raw_lower = (str(cell("pupil")) + " " + note + " " + paid).lower()
    # status
    if pd.isna(fees) and ("not attending" in raw_lower or "not in school" in raw_lower):
        status = "Not attending"
    elif pd.isna(fees):
        status = "Not invoiced"
    else:
        status = "Enrolled"
    cls = "Wednesday" if "club" in note.lower() and status == "Enrolled" else "Saturday"
    # per-child facts
    sess_override, override_note = None, ""
    m = re.search(r"(\d+)\s*sessions", raw_lower)
    if m: sess_override = float(m.group(1)); override_note = f"roster note: {m.group(1)} sessions"
    discount = 0.0; sup_pos = 0.0
    if pd.notna(sup):
        if sup < 0: discount = round(-float(sup), 2)
        else: sup_pos = float(sup)
    rows.append(dict(sur=sur, first=first, inv=inv, inv_raw=inv_raw, status=status, cls=cls,
                     sess_override=sess_override, override_note=override_note,
                     discount=discount, reg_flag="Y" if (pd.notna(reg) and reg > 0) else "N",
                     sup_flag="Y" if sup_pos > 0 else "N",
                     r_fees=0.0 if pd.isna(fees) else float(fees), r_sup=sup_pos,
                     r_reg=0.0 if pd.isna(reg) else float(reg), r_disc=discount,
                     paid=paid, note=note, inferred=inferred, raw=str(cell("pupil"))))
ch = pd.DataFrame(rows)

# An inferred surname yields to the roster's own evidence. "Martin Lucas" on invoice 2026-999,
# where MARTIN Ines and MARTIN Jules already sit, is a third Martin child, not a Lucas: a token
# of the name equals the surname of a properly written row on the same invoice. That is two
# independent facts agreeing, so the inferred reading is replaced and the family stays whole.
known = ch[~ch.inferred & (ch.inv != "")].groupby("inv")["sur"].agg(set)
for i in ch.index[ch.inferred & (ch.inv != "")]:
    toks = {norm(t).upper() for t in re.split(r"[\s-]+", ch.at[i, "raw"])}
    match = [s_ for s_ in known.get(ch.at[i, "inv"], set()) if norm(s_).upper() in toks]
    if len(match) == 1 and match[0] != ch.at[i, "sur"]:
        first_tok = next(t for t in ch.at[i, "raw"].split() if norm(t).upper() != norm(match[0]).upper())
        ch.at[i, "note"] = ch.at[i, "note"].replace(f"read as surname '{ch.at[i, 'sur']}'; confirm",
                                                    f"read as surname '{match[0]}', shared invoice with that family")
        ch.at[i, "sur"], ch.at[i, "first"] = match[0], re.sub(r"[^A-Za-zÀ-ÿ'\-]", "", first_tok)
n_late = int(ch.inferred.sum()) if len(ch) else 0
ch = ch.drop(columns=["inferred", "raw"])

# ---------------------------------------------------------------- families
def same_family(a, b):
    a, b = norm(a), norm(b)
    return a[:4] == b[:4] or SequenceMatcher(None, a, b).ratio() >= 0.8

fam_of = {}
groups = defaultdict(list)          # key -> list of clusters (each cluster = list of row idx)
for idx, r in ch.iterrows():
    key = r.inv if r.inv else "NOINV"
    placed = False
    for cl in groups[key]:
        if same_family(ch.loc[cl[0], "sur"], r.sur):
            cl.append(idx); placed = True; break
    if not placed: groups[key].append([idx])
clusters = [cl for cls_ in groups.values() for cl in cls_]
clusters.sort(key=lambda cl: (norm(ch.loc[cl[0], "sur"]), ch.loc[cl[0], "inv"]))
fam_rows = []
for n, cl in enumerate(clusters, 1):
    fid = f"FAM-{n:03d}"
    surs = sorted({ch.loc[i, "sur"] for i in cl}, key=len, reverse=True)
    invs = sorted({ch.loc[i, "inv"] for i in cl if ch.loc[i, "inv"]})
    for i in cl: fam_of[i] = fid
    shared = [k for k in invs if len(groups[k]) > 1]
    fam_rows.append(dict(fid=fid, name=" / ".join(surs), invs=", ".join(invs),
                         note=("invoice number shared with another family  -  " + ", ".join(shared)) if shared else ""))
ch["fid"] = ch.index.map(fam_of)
fam = pd.DataFrame(fam_rows)
ch = ch.sort_values(["fid", "sur", "first"]).reset_index(drop=True)
ch["cid"] = [f"CH-{i+1:03d}" for i in range(len(ch))]
# ---------------------------------------------------------------- bank
b = pd.read_excel(BANK, header=None)
# The fixture has no header and a fixed column order. Barclays writes a header row and orders the
# columns differently, so when a header is there the columns are taken by name and never by index.
head = [str(x).strip() for x in b.iloc[0]] if len(b) else []
if len(b) and bank_header(b.iloc[0]):
    at = {f: next((head.index(n) for n in names_ if n in head), None) for f, names_ in BANK_COLS.items()}
    missing = [f for f, i in at.items() if i is None and f != "cat"]
    if missing:
        sys.exit(f"{BANK}: bank export has no {', '.join(missing)} column. Headings found: {head}")
    body = b.iloc[1:].reset_index(drop=True)
    b = pd.DataFrame({f: (body[i] if i is not None else None) for f, i in at.items()})
    b["note"] = None; b["extra"] = None
else:
    b = b.reindex(columns=range(len(BANK_POSITIONAL)))
    b.columns = BANK_POSITIONAL
# Real exports carry blank spacer rows. Without this a blank amount reaches the `r.amount < 0`
# test below as NaT and raises "'<' not supported between instances of 'NaTType' and 'int'".
b = b.dropna(subset=["date", "amount"], how="all").reset_index(drop=True)
# A UK bank writes dd/mm/yyyy, and an export that has been through Excel can carry its amounts as
# text ("1,234.50"). Neither may cost a receipt silently: a line with a memo but no readable date
# or amount stops the build, so the office fixes the file rather than losing the money.
b["date"] = pd.to_datetime(b["date"], errors="coerce", dayfirst=True)
b["amount"] = pd.to_numeric(b["amount"].astype(str).str.replace(r"[£,\s]", "", regex=True), errors="coerce")
unreadable = b[(b["date"].isna() | b["amount"].isna()) & b["memo"].notna()]
if len(unreadable):
    sys.exit(f"{BANK}: {len(unreadable)} line(s) with a memo but no readable date or amount, e.g. "
             f"{unreadable.iloc[0].to_dict()}. Fix the export; nothing was written.")
b = b[b["date"].notna() & b["amount"].notna()].reset_index(drop=True)
# The real statement is the whole account, not just fee income: rent, utilities and card spending
# share it with the receipts. Drop outgoing lines in the configured categories so the office is not
# asked to review the electricity bill. Only negatives qualify, so no receipt can go missing here.
cat_l = b["cat"].astype(str).str.strip().str.lower()
not_fee = (b["amount"] < 0) & cat_l.isin(NOT_A_RECEIPT)
n_not_fee = int(not_fee.sum())
b = b[~not_fee].reset_index(drop=True)
CODE = re.compile(r"^(FT|BGC|BG|BBP|BP|B)$")
# Reference spellings, built from the chosen term's year rather than written out: 2026-047 and
# 2026047 for the current term, 2025-6xx and the two-digit forms parents use for the prior one.
# engine.py builds the same patterns from its own Config; this copy exists because build_ledger.py
# only needs to decide which term a receipt belongs to, never which family.
CUR_REF_RE = re.compile(rf"{YEAR_CUR}\s*-?\s*(\d{{3}})")
CUR_REF_SHORT_RE = re.compile(rf"(?<!\d){YEAR_CUR[2:]}(\d{{3}})(?!\d)")
PRIOR_REF_RE = re.compile(rf"{YEAR_PRIOR}\s*-?\s*(\d{{3}})|(?<!\d)(2[1-4])\s*-?\s*(1\d\d)(?!\d)|(?<!\d)(2[1-4])(\d{{3}})(?!\d)")
CUR_WINDOW_FROM = pd.Timestamp(CUR["start"])
rec = []
# Current-term invoice totals, taken from the roster: each enrolled invoice's family total, half of it
# (one sibling's share) and each enrolled child's own line. Used only to notice a prior-term
# reference sitting on a payment whose amount and date say current term. Deciding which family a
# receipt belongs to is engine.py's job.
enr = ch[ch.status == "Enrolled"]
child_tot = (enr.r_fees + enr.r_sup + enr.r_reg - enr.r_disc).round(2)
# grouped by family as well as invoice: an invoice number shared by two families is two invoices
inv_tot = child_tot.groupby([enr.fid, enr.inv]).sum().drop("", level=1, errors="ignore")
cur_amounts = {float(x) for x in (*inv_tot, *(inv_tot / 2), *child_tot)}
def is_cur_amount(amount, amounts):
    """Decide whether a receipt amount equals one of the roster figures to the penny; a half-share's odd half-penny may have been rounded either way."""
    return any(abs(amount - a) < 0.006 for a in amounts)
for _, r in b.iterrows():
    memo = str(r.memo)
    payer = memo[:23].strip()
    tail = memo[23:].strip().split()
    code = tail[-1] if tail and CODE.match(tail[-1]) else ""
    ref = " ".join(tail[:-1] if code else tail)
    M = norm(memo).upper()
    # A reference is this term's only if its year matches AND its number is inside this term's
    # series; where no series is set the year decides on its own, as it always did.
    m = next((mm for mm in CUR_REF_RE.finditer(M) if in_series(mm.group(1), CUR_SERIES)), None) \
        or next((mm for mm in CUR_REF_SHORT_RE.finditer(M) if in_series(mm.group(1), CUR_SERIES)), None)
    cur_ref = f"{YEAR_CUR}-{m.group(1)}" if m else ""
    pm = next((mm for mm in PRIOR_REF_RE.finditer(M) if not mm.group(1) or in_series(mm.group(1), PRIOR_SERIES)), None)
    prior_ref = pm.group(0) if (pm and not cur_ref) else ""
    # Which term a receipt belongs to is a property of the receipt, so it is settled here. Which
    # family it belongs to is a judgement on ranked evidence, and that lives in engine.py.
    if cur_ref: term = TERM_CUR
    elif prior_ref: term = TERM_PRIOR
    else: term = TERM_PRIOR if r.date < CUR_WINDOW_FROM else TERM_CUR
    flags = []
    if r.amount < 0: flags.append("outflow / refund")
    if len(memo) >= 44: flags.append("bank truncated the reference field")
    if prior_ref and r.date >= CUR_WINDOW_FROM and is_cur_amount(float(r.amount), cur_amounts):
        flags.append(f"prior-term ref but {TERM_CUR} amount and date  -  parent may have reused old reference")
    extra = " | ".join(str(x).strip() for x in [r.note, r.extra] if pd.notna(x))
    rec.append(dict(date=r.date.date(), amount=float(r.amount), payer=payer, ref=ref, memo=memo,
                    term=term, flag="; ".join(flags), extra=extra))
rc = pd.DataFrame(rec).sort_values("date").reset_index(drop=True)

# ---------------------------------------------------------------- workbook
wb = Workbook()
F = "Arial"
H_FILL = PatternFill("solid", fgColor="D9D9D9")
Y_FILL = PatternFill("solid", fgColor="FFFF00")
BLUE, BLACK, GREEN = Font(name=F, size=10, color="0000FF"), Font(name=F, size=10), Font(name=F, size=10, color="008000")
BOLD = Font(name=F, size=10, bold=True)
TITLE = Font(name=F, size=13, bold=True)
GBP = '£#,##0.00;(£#,##0.00);-'
thin = Side(style="thin", color="BFBFBF")

def header(ws, row, cols):
    for j, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=j, value=h); c.font = BOLD; c.fill = H_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
def widths(ws, w):
    for j, x in enumerate(w, 1): ws.column_dimensions[get_column_letter(j)].width = x
def put(ws, row, col, val, font=BLACK, fmt=None, fill=None):
    c = ws.cell(row=row, column=col, value=val); c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    return c

# ---- README
ws = wb.active; ws.title = "README"
shared_inv = sorted({k for k, v in groups.items() if k != "NOINV" and len(v) > 1})
n_noinv = int((ch.inv == "").sum()); n_disc = int((ch.discount > 0).sum())
def seq_num(inv):
    """The three-digit part of an invoice number, or None if it is not spelled like one."""
    m = re.match(r"^\d{4}-(\d{1,3})$", str(inv))
    return int(m.group(1)) if m else None
# Out of sequence means outside the series this term issues, which is a term fact, not a constant.
lo, hi = CUR_SERIES or (0, 200)
n_out_of_seq = sorted({i for i in ch.inv if i and (seq_num(i) is None or not lo <= seq_num(i) <= hi)})
n_noinv_paid = int(((ch.inv == "") & (ch.paid != "")).sum())
def plural(n, one, many):
    """Decide the word form that agrees with a count: `one` when n == 1, else `many`."""
    return one if n == 1 else many
noinv_line = f"{n_noinv} enrolled {plural(n_noinv, 'child has', 'children have')} no invoice number and no fee."
if n_noinv_paid:
    noinv_line += (f" {'That child' if n_noinv == 1 else f'{n_noinv_paid} of them'} {plural(n_noinv_paid, 'has', 'have')}"
                   " a payment note on the roster, so money may have arrived with nothing to match it to.")
def rules_disagree(r):
    """Decide whether the rule-based expected total differs from what the roster invoiced by 0.02 or more."""
    exp_rule = round(r.r_fees - r.discount + (25.0 if r.reg_flag == "Y" else 0) + (25.0 if r.sup_flag == "Y" else 0), 2)
    roster_tot = round(r.r_fees + r.r_sup + r.r_reg - r.r_disc, 2)
    return abs(exp_rule - roster_tot) >= 0.02
# counted once for the README, reused to fill the same rows red on Children
mismatched = {i for i, r in ch.iterrows() if r.status == "Enrolled" and rules_disagree(r)}
lines = [
 ("Fee ledger  -  how it works", TITLE),
 ("Built from the pupil roster and the bank export. Health, guardian-contact and free-text pupil columns were deliberately not carried over.", BLACK),
 ("", BLACK),
 ("Where to look", BOLD),
 ("Position  -  the answer sheet. One row per family: what they owe, what came in, every receipt reference, whether it is settled. Colour-coded. Start here.", BLACK),
 ("Review  -  the only sheet that needs you. Receipts the engine would not allocate, with candidates and a reason. Decide, type the family ID into Receipts column F, re-run.", BLACK),
 ("Receipts  -  every bank line and what the engine did with it.", BLACK),
 ("Children / Families  -  the roster, and the fee each child should be charged. Difference and Check compare that with what the roster actually invoiced.", BLACK),
 ("Rates / Terms  -  the fee rules and the session counts. Change a number here and everything recalculates.", BLACK),
 ("Summary  -  headline figures.", BLACK),
 ("", BLACK),
 ("Colour key", BOLD),
 ("Blue text = typed in from a source file. Black = formula. Yellow fill = for you to fill in. Green = written by the engine.", BLACK),
 ("Green row / red row on Position = settled / owing. Amber = needs a look before you chase: a receipt sitting in Review, or money tagged to the wrong term.", BLACK),
 ("Red cell in Children = the roster invoiced something the fee rules do not produce.", BLACK),
 ("", BLACK),
 ("How a receipt gets a family", BOLD),
 ("Ranked evidence, first hit wins: your manual override; then the invoice reference, or a payer name already seen with a verified reference; then a unique surname in the memo; then fuzzy evidence such as a truncated surname or a child's first name. The first three allocate automatically, the rest go to Review. If the reference and the payer name disagree, nothing is allocated and the line is flagged.", BLACK),
 ("", BLACK),
 ("Assumptions  -  confirm these with the office", BOLD),
 ("Rates come from the January price list: £195 / £345 / £460 / £571 per 10 sessions for 1-4 siblings on Saturdays, £160 per child for Wednesdays, £25 registration, £25 supplies.", BLACK),
 (f"Sessions per term are inferred from the amounts invoiced, not from a timetable: {CUR['sat']} Saturdays this term, {PRIOR['sat']} last term, {CUR['wed']} Wednesdays. Sibling pricing is treated as a family rate split evenly between the children.", BLACK),
 (f"The {TERM_PRIOR} invoice register was never supplied, so {TERM_PRIOR} receipts can be tied to a family but not to an invoice." if PRIOR["loaded"] != "Yes" else f"The {TERM_PRIOR} invoice register is loaded, so prior-term receipts can be reconciled against an invoice.", BLACK),
 ("", BLACK),
 ("What the roster got wrong", BOLD),
 (f"Invoice numbers shared by unrelated families: {', '.join(shared_inv)}  -  so the invoice number cannot be the family key, hence FAM-nnn.", BLACK),
 (noinv_line, BLACK),
 (f"{n_disc} discounts were stored as negative numbers in the supplies column; they are a proper Discount column here. Invoice numbers out of sequence: {', '.join(n_out_of_seq)}.", BLACK),
 (f"{len(mismatched)} {plural(len(mismatched), 'invoice', 'invoices')} whose add-ons or discount the fee rules cannot reproduce. The Check column on Children also recomputes tuition from sessions and sibling count, so it can flag rows this count does not.", BLACK),
 ("", BLACK),
 ("To re-run", BOLD),
 ("Replace the bank export, run build_ledger.py then reconcile.py. Copy column F of Receipts and the alias column of Families first  -  manual decisions live in the workbook, not in the scripts.", BLACK),
]
for i, (t, f) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=t); c.font = f; c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 150

# ---- Rates
ws = wb.create_sheet("Rates")
put(ws, 1, 1, "Fee rules  -  inputs (blue). Source: school price list, January 2026.", TITLE)
header(ws, 3, ["Item", "Basis", "£ per session / per item", "Source / note"])
rates = [
 (4, "Saturday school  -  1 child", "per family per session", 19.50, "£195 per 10 sessions"),
 (5, "Saturday school  -  2 siblings", "per family per session", 34.50, "£345 per 10 sessions"),
 (6, "Saturday school  -  3 siblings", "per family per session", 46.00, "£460 per 10 sessions"),
 (7, "Saturday school  -  4 siblings", "per family per session", 57.10, "£571 per 10 sessions"),
 (9, "Wednesday class  -  per child", "per child per session", 16.00, "£160 per 10 sessions; no sibling tier listed"),
 (11, "Registration fee", "per child, one-off at first registration", 25.00, "Price list says per family for Wednesday  -  treated per child here (one Wednesday child only)"),
 (12, "Supplies / annual membership", "per child per year", 25.00, "Price list"),
 (14, "Trial  -  Saturday, 1 child", "information only, not used in formulas", 19.50, "£34.50 / £46 / £57 for 2 / 3 / 4 siblings"),
 (15, "Trial  -  Wednesday", "information only, not used in formulas", 16.00, ""),
]
for row, item, basis, val, src in rates:
    put(ws, row, 1, item); put(ws, row, 2, basis); put(ws, row, 3, val, BLUE, GBP); put(ws, row, 4, src)
widths(ws, [34, 40, 22, 70]); ws.freeze_panes = "A4"

# ---- Terms
ws = wb.create_sheet("Terms")
put(ws, 1, 1, "Invoicing terms  -  inputs (blue). Add a row per new term; the register for each term goes in Invoices.", TITLE)
header(ws, 3, ["Term ID", "Description", "Saturday sessions", "Wednesday sessions", "Invoice series", "Register loaded?", "Receipts window from", "Receipts window to", "Note"])
# Row 4 is the prior term and row 5 the current one. The Children sheet's session lookup and
# reconcile.py's ageing formula both address those two rows, so the order is load-bearing.
terms = [(4 + i, t["id"], t["desc"], t["sat"], t["wed"], t["series"], t["loaded"], t["start"], t["end"], t["note"])
         for i, t in enumerate((PRIOR, CUR))]
for row, *vals in terms:
    for j, v in enumerate(vals, 1):
        put(ws, row, j, v, BLUE if j in (3, 4, 5, 6, 7, 8) else BLACK, "dd/mm/yyyy" if j in (7, 8) else None)
widths(ws, [12, 28, 12, 12, 20, 12, 14, 14, 90]); ws.freeze_panes = "A4"

# ---- Families
ws = wb.create_sheet("Families")
put(ws, 1, 1, "Families  -  stable key. One family can have several invoice numbers over time; the family ID never changes.", TITLE)
header(ws, 3, ["Family ID", "Family surname(s)", "Children on roster", f"Children billed ({TERM_CUR})", f"Invoice no(s) {TERM_CUR}", "Bank payer alias (fill in; separate with ;)", "Payer names seen with a verified reference (engine)", "Note"])
for i, r in fam.iterrows():
    row = 4 + i
    put(ws, row, 1, r.fid, BLUE); put(ws, row, 2, r["name"], BLUE)
    put(ws, row, 3, f'=COUNTIF(Children!$B$4:$B$500,$A{row})')
    put(ws, row, 4, f'=COUNTIFS(Children!$B$4:$B$500,$A{row},Children!$F$4:$F$500,"Enrolled")')
    put(ws, row, 5, r.invs, BLUE); put(ws, row, 6, "", fill=Y_FILL); put(ws, row, 7, ""); put(ws, row, 8, r.note, BLUE)
widths(ws, [11, 32, 10, 12, 22, 30, 40, 60]); ws.freeze_panes = "A4"; ws.auto_filter.ref = f"A3:H{3+len(fam)}"

# ---- Children
ws = wb.create_sheet("Children")
put(ws, 1, 1, "Children  -  facts per child (blue) drive the expected fee (black). 'Check' must be OK; anything else is an invoice that does not follow the rules.", TITLE)
cols = ["Child ID", "Family ID", "Surname", "First name", "Class", "Status", "Term", "Invoice no (roster)",
        "Sessions", "Sessions override", "Siblings billed", "Rate per session", "Tuition", "Discount", "Add-ons charged",
        "Reg fee", "Supplies", "Expected total", "Invoiced on roster", "Difference", "Check", "Note"]
header(ws, 3, cols)
first, last = 4, 3 + len(ch)
for i, r in ch.iterrows():
    row = first + i
    for j, v in enumerate([r.cid, r.fid, r.sur, r["first"], r.cls, r.status, TERM_CUR, r.inv_raw], 1):
        put(ws, row, j, v, BLUE)
    put(ws, row, 9, f'=IF($J{row}<>"",$J{row},IF($E{row}="Wednesday",INDEX(Terms!$D$4:$D$5,MATCH($G{row},Terms!$A$4:$A$5,0)),INDEX(Terms!$C$4:$C$5,MATCH($G{row},Terms!$A$4:$A$5,0))))', fmt="0.0")
    put(ws, row, 10, r.sess_override if r.sess_override else None, BLUE, "0.0")
    put(ws, row, 11, f'=COUNTIFS($B${first}:$B$500,$B{row},$F${first}:$F$500,"Enrolled")')
    put(ws, row, 12, f'=IF($F{row}<>"Enrolled",0,IF($K{row}=0,0,IF($E{row}="Wednesday",Rates!$C$9,INDEX(Rates!$C$4:$C$7,MIN($K{row},4))/$K{row})))', fmt=GBP)
    put(ws, row, 13, f'=ROUND($L{row}*$I{row},2)', fmt=GBP)
    put(ws, row, 14, r.discount, BLUE, GBP)
    addons = ", ".join(x for x, f in (("reg", r.reg_flag == "Y"), ("supplies", r.sup_flag == "Y")) if f)
    put(ws, row, 15, addons, BLUE)
    put(ws, row, 16, f'=IF(ISNUMBER(SEARCH("reg",$O{row})),Rates!$C$11,0)', fmt=GBP)
    put(ws, row, 17, f'=IF(ISNUMBER(SEARCH("supplies",$O{row})),Rates!$C$12,0)', fmt=GBP)
    put(ws, row, 18, f'=$M{row}-$N{row}+$P{row}+$Q{row}', fmt=GBP)
    put(ws, row, 19, round(r.r_fees + r.r_sup + r.r_reg - r.r_disc, 2), BLUE, GBP)
    put(ws, row, 20, f'=ROUND($R{row}-$S{row},2)', fmt=GBP)
    put(ws, row, 21, f'=IF(ABS($T{row})<0.02,"OK","CHECK")')
    note = " | ".join(x for x in [r.override_note, r.note, ("roster PAID: " + r.paid) if r.paid else ""] if x)
    put(ws, row, 22, note, BLUE)
    if i in mismatched:
        for c in range(1, 23): ws.cell(row, c).fill = PatternFill("solid", fgColor="FCE4E4")
widths(ws, [9, 10, 20, 14, 10, 13, 10, 12, 8, 9, 8, 10, 10, 9, 14, 8, 9, 11, 11, 10, 8, 55])
ws.freeze_panes = "E4"; ws.auto_filter.ref = f"A3:V{last}"
ws["J3"].comment = Comment("Leave blank for a standard term. Enter sessions only when the child is billed for a different number, e.g. 9 sessions, or 10.5 for half a sibling share.", "ledger")
ws.conditional_formatting.add(f"A4:V{last}", FormulaRule(formula=[f'$U4="CHECK"'], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=F, size=10, color="9C0006")))

# ---- Receipts
ws = wb.create_sheet("Receipts")
put(ws, 1, 1, "Receipts  -  one row per bank line. Blue = from the bank file. Yellow = your decision. Everything else is written by the engine.", TITLE)
# Columns 9-13 are written by reconcile.py; column 10 is what $G and $J below read back.
header(ws, 3, ["Date", "Amount", "Payer (bank)", "Reference (bank)", "Term", "Family override (fill in)", "Family",
               "Allocated", "Tier", "Family (engine)", "Candidates", "Why", "Amount check", "Flags", "Full memo"])
for i, r in rc.iterrows():
    row = 4 + i
    put(ws, row, 1, r.date, BLUE, "dd/mm/yyyy"); put(ws, row, 2, r.amount, BLUE, GBP)
    put(ws, row, 3, r.payer, BLUE); put(ws, row, 4, r.ref, BLUE); put(ws, row, 5, r.term, BLUE)
    put(ws, row, 6, "", fill=Y_FILL)
    put(ws, row, 7, f'=IF($F{row}<>"",$F{row},$J{row})')
    put(ws, row, 8, f'=IF($G{row}<>"","yes","")')
    put(ws, row, 14, r.flag or r.extra, BLUE); put(ws, row, 15, r.memo, BLUE)
widths(ws, [11, 11, 22, 24, 10, 14, 11, 9, 6, 11, 16, 60, 20, 40, 46])
ws.freeze_panes = "C4"; ws.auto_filter.ref = f"A3:O{3 + len(rc)}"

wb.save(OUT)
print("roster sheet", repr(ROSTER_SHEET), "| headings on row", ROSTER_HEADER + 1)
print("saved", OUT, "| children", len(ch), "| families", len(fam), "| receipts", len(rc))
print("statuses", ch.status.value_counts().to_dict(), "| rows not in SURNAME-first form", n_late)
if n_not_fee:
    print(f"dropped {n_not_fee} outgoing non-fee line(s): {', '.join(sorted(NOT_A_RECEIPT))}")
print("shared invoices", shared_inv, "| no-invoice children", n_noinv, "| discounts", n_disc, "| out of seq", n_out_of_seq)
