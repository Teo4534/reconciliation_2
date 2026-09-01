"""
build_ledger.py — turn a pupil roster and a bank export into a structured fee ledger.

    python build_ledger.py <roster.xlsx> <bank.xlsx> <ledger_out.xlsx>

Writes Rates, Terms, Families, Children and Receipts. The fee rules live in Rates as data,
so a price change is a cell edit, not a code change. Run reconcile.py next to allocate receipts.
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

ROSTER = sys.argv[1] if len(sys.argv) > 1 else "examples/roster.xlsx"
BANK = sys.argv[2] if len(sys.argv) > 2 else "examples/bank.xlsx"
OUT = sys.argv[3] if len(sys.argv) > 3 else "examples/fee_ledger.xlsx"
TERM_CUR, TERM_PRIOR = "JAN-2026", "AUT-2025"

def norm(x):
    return unicodedata.normalize("NFKD", str(x)).encode("ascii", "ignore").decode()

# ---------------------------------------------------------------- roster
s = pd.read_excel(ROSTER)
s.columns = [str(c).strip() for c in s.columns]
s = s[s["Statut:"] == "Inscrit"].copy().reset_index(drop=True)
note_cols = [c for c in s.columns if c.startswith("Unnamed: 1")]

def parse_name(raw):
    t = re.sub(r"\s*-\s*", "-", str(raw)).strip()
    t = t.split("(")[0]
    toks = t.split()
    sur = []
    for tok in toks:
        if tok.isupper() and len(tok) > 1: sur.append(tok)
        else: break
    if not sur:
        sur, toks = [toks[0]], toks
        first = toks[1] if len(toks) > 1 else ""
    else:
        rest = toks[len(sur):]
        first = rest[0] if rest else ""
    first = re.sub(r"[^A-Za-zÀ-ÿ'\-]", "", first)
    return " ".join(sur).upper(), first

rows = []
for i, r in s.iterrows():
    sur, first = parse_name(r["LES ELEVES"])
    inv_raw = "" if pd.isna(r["INVOICE NUMBER"]) else str(r["INVOICE NUMBER"]).strip()
    inv = re.sub(r"[a.]$", "", inv_raw)
    note = " | ".join(str(r[c]).strip() for c in note_cols if pd.notna(r[c]))
    paid = "" if pd.isna(r["PAID"]) else str(r["PAID"]).strip()
    fees = r["FEES"]; sup = r["OFFICE SUPPLIES"]; reg = r["REG FEES ONE OFF"]
    raw_lower = (str(r["LES ELEVES"]) + " " + note + " " + paid).lower()
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
                     paid=paid, note=note))
ch = pd.DataFrame(rows)

# hand-resolved overrides where the roster amount implies a non-standard session count
def set_override(mask, sessions, why):
    ch.loc[mask, "sess_override"] = sessions
    ch.loc[mask, "override_note"] = why
set_override((ch.r_fees == 181.13), 10.5, "half of sibling share (£181.13) — joined mid-term? confirm")
set_override((ch.r_fees == 546.0), 28, "£546 = 28 × £19.50 — unexplained, confirm with office")

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
                         note=("invoice number shared with another family — " + ", ".join(shared)) if shared else ""))
ch["fid"] = ch.index.map(fam_of)
fam = pd.DataFrame(fam_rows)
ch = ch.sort_values(["fid", "sur", "first"]).reset_index(drop=True)
ch["cid"] = [f"CH-{i+1:03d}" for i in range(len(ch))]
inv_to_fams = defaultdict(set)
for _, r in ch.iterrows():
    if r.inv: inv_to_fams[r.inv].add(r.fid)

# ---------------------------------------------------------------- bank
b = pd.read_excel(BANK, header=None)
names = ["date", "amount", "memo", "cat", "note", "extra"]
b = b.reindex(columns=range(len(names)))      # tolerate trailing empty columns dropped on read
b.columns = names
CODE = re.compile(r"^(FT|BGC|BG|BBP|BP|B)$")
rec = []
cur_amounts = {409.5, 362.25, 724.5, 322.0, 966.0, 434.5, 459.5, 749.5, 774.5, 387.25, 372.25}
# surname index for memo scanning (whole-word, 4+ letters; hyphenated names split into parts)
sur_index = defaultdict(set)
for _, r in ch.iterrows():
    for part in re.split(r"[-\s]", norm(r.sur).upper()):
        if len(part) >= 4: sur_index[part].add(r.fid)
def fams_in_memo(text):
    T = norm(text).upper()
    hits = set()
    for part, fids in sur_index.items():
        if re.search(r"(?<![A-Z])" + re.escape(part) + r"(?![A-Z])", T): hits |= fids
    return hits
for _, r in b.iterrows():
    memo = str(r.memo)
    payer = memo[:23].strip()
    tail = memo[23:].strip().split()
    code = tail[-1] if tail and CODE.match(tail[-1]) else ""
    ref = " ".join(tail[:-1] if code else tail)
    M = norm(memo).upper()
    m = re.search(r"2026\s*-?\s*(\d{3})", M) or re.search(r"(?<!\d)26(\d{3})(?!\d)", M)
    cur_ref = f"2026-{m.group(1)}" if m else ""
    pm = re.search(r"2025\s*-?\s*(\d{3})|(?<!\d)(2[1-4])\s*-?\s*(1\d\d)(?!\d)|(?<!\d)(2[1-4])(\d{3})(?!\d)", M)
    prior_ref = pm.group(0) if (pm and not cur_ref) else ""
    bm = re.search(r"(?<!\d)(\d{3})(?!\d)", M) if not (cur_ref or prior_ref) else None
    bare = f"2026-{bm.group(1)}" if bm else ""
    if cur_ref: term, basis = TERM_CUR, "current-term reference"
    elif prior_ref: term, basis = TERM_PRIOR, "prior-term reference"
    else:
        term = TERM_PRIOR if r.date < pd.Timestamp("2025-12-15") else TERM_CUR
        basis = "date only — no usable reference"
    auto, how, suggest = "", "", ""
    named = fams_in_memo(memo)
    if cur_ref:
        f = inv_to_fams.get(cur_ref, set())
        if len(f) == 1:
            cand = next(iter(f))
            if named and cand not in named:
                how = f"CONFLICT — reference {cur_ref} points to {cand} but the memo names " + ", ".join(sorted(named)) + " (parent may have typed the wrong invoice number)"
                if len(named) == 1: suggest = next(iter(named))
            else:
                auto, how = cand, "invoice reference"
        elif len(f) > 1:
            inter = f & named
            if len(inter) == 1: auto, how = next(iter(inter)), "invoice reference (shared number resolved by surname in memo)"
            else: how = "ambiguous — invoice number shared by " + ", ".join(sorted(f))
        else: how = "reference not in roster"
    elif prior_ref:
        how = "prior-term register not loaded"
        if len(named) == 1: suggest = next(iter(named))
    elif bare and len(inv_to_fams.get(bare, set())) == 1:
        suggest, how = next(iter(inv_to_fams[bare])), "bare 3-digit ref — series assumed, review"
    elif len(named) == 1:
        suggest, how = next(iter(named)), "no usable reference — surname found in memo, review"
    elif len(named) > 1:
        how = "no usable reference — memo names several families: " + ", ".join(sorted(named))
    else: how = "no usable reference — name match needed (project 2)"
    flags = []
    if r.amount < 0: flags.append("outflow / refund")
    if len(memo) >= 44: flags.append("bank truncated the reference field")
    if prior_ref and r.date >= pd.Timestamp("2026-01-01") and round(float(r.amount), 2) in cur_amounts:
        flags.append("prior-term ref but 2026-term amount and date — parent may have reused old reference")
    extra = " | ".join(str(x).strip() for x in [r.note, r.extra] if pd.notna(x))
    rec.append(dict(date=r.date.date(), amount=float(r.amount), payer=payer, ref=ref, code=code, memo=memo,
                    cur_ref=cur_ref, prior_ref=prior_ref, term=term, basis=basis, auto=auto, how=how,
                    suggest=suggest, flag="; ".join(flags), extra=extra))
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
n_out_of_seq = sorted({i for i in ch.inv if i and int(i.split("-")[1]) > 200})
lines = [
 ("Fee ledger — how it works", TITLE),
 ("Built from the pupil roster and the bank export. Health, guardian-contact and free-text pupil columns were deliberately not carried over.", BLACK),
 ("", BLACK),
 ("Where to look", BOLD),
 ("Position — the answer sheet. One row per family: what they owe, what came in, every receipt reference, whether it is settled. Colour-coded. Start here.", BLACK),
 ("Review — the only sheet that needs you. Receipts the engine would not allocate, with candidates and a reason. Decide, type the family ID into Receipts column F, re-run.", BLACK),
 ("Receipts — every bank line and what the engine did with it.", BLACK),
 ("Children / Families — the roster, and the fee each child should be charged. Difference and Check compare that with what the roster actually invoiced.", BLACK),
 ("Rates / Terms — the fee rules and the session counts. Change a number here and everything recalculates.", BLACK),
 ("Summary — headline figures.", BLACK),
 ("", BLACK),
 ("Colour key", BOLD),
 ("Blue text = typed in from a source file. Black = formula. Yellow fill = for you to fill in. Green = written by the engine.", BLACK),
 ("Green row / red row on Position = settled / owing. Amber = needs a look before you chase: a receipt sitting in Review, or money tagged to the wrong term.", BLACK),
 ("Red cell in Children = the roster invoiced something the fee rules do not produce.", BLACK),
 ("", BLACK),
 ("How a receipt gets a family", BOLD),
 ("Ranked evidence, first hit wins: your manual override; then the invoice reference, or a payer name already seen with a verified reference; then a unique surname in the memo; then fuzzy evidence such as a truncated surname or a child's first name. The first three allocate automatically, the rest go to Review. If the reference and the payer name disagree, nothing is allocated and the line is flagged.", BLACK),
 ("", BLACK),
 ("Assumptions — confirm these with the office", BOLD),
 ("Rates come from the January price list: £195 / £345 / £460 / £571 per 10 sessions for 1-4 siblings on Saturdays, £160 per child for Wednesdays, £25 registration, £25 supplies.", BLACK),
 ("Sessions per term are inferred from the amounts invoiced, not from a timetable: 21 Saturdays this term, 11 last term, 20 Wednesdays. Sibling pricing is treated as a family rate split evenly between the children.", BLACK),
 ("The autumn invoice register was never supplied, so autumn receipts can be tied to a family but not to an invoice.", BLACK),
 ("", BLACK),
 ("What the roster got wrong", BOLD),
 (f"Invoice numbers shared by unrelated families: {', '.join(shared_inv)} — so the invoice number cannot be the family key, hence FAM-nnn.", BLACK),
 (f"{n_noinv} enrolled children have no invoice number and no fee. Two of them have payment notes, so money may have arrived with nothing to match it to.", BLACK),
 (f"{n_disc} discounts were stored as negative numbers in the supplies column; they are a proper Discount column here. Invoice numbers out of sequence: {', '.join(n_out_of_seq)}.", BLACK),
 ("One invoice charges £18 supplies instead of £25 — the only invoice the fee rules cannot reproduce.", BLACK),
 ("", BLACK),
 ("To re-run", BOLD),
 ("Replace the bank export, run build_ledger.py then reconcile.py. Copy column F of Receipts and the alias column of Families first — manual decisions live in the workbook, not in the scripts.", BLACK),
]
for i, (t, f) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=t); c.font = f; c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 150

# ---- Rates
ws = wb.create_sheet("Rates")
put(ws, 1, 1, "Fee rules — inputs (blue). Source: charity price list, January 2026 (photo supplied by user).", TITLE)
header(ws, 3, ["Item", "Basis", "£ per session / per item", "Source / note"])
rates = [
 (4, "Saturday school — 1 child", "per family per session", 19.50, "£195 per 10 sessions"),
 (5, "Saturday school — 2 siblings", "per family per session", 34.50, "£345 per 10 sessions"),
 (6, "Saturday school — 3 siblings", "per family per session", 46.00, "£460 per 10 sessions"),
 (7, "Saturday school — 4 siblings", "per family per session", 57.10, "£571 per 10 sessions"),
 (9, "Wednesday class — per child", "per child per session", 16.00, "£160 per 10 sessions; no sibling tier listed"),
 (11, "Registration fee", "per child, one-off at first registration", 25.00, "Price list says per family for Wednesday — treated per child here (one Wednesday child only)"),
 (12, "Supplies / annual membership", "per child per year", 25.00, "Price list"),
 (14, "Trial — Saturday, 1 child", "information only, not used in formulas", 19.50, "£34.50 / £46 / £57 for 2 / 3 / 4 siblings"),
 (15, "Trial — Wednesday", "information only, not used in formulas", 16.00, ""),
]
for row, item, basis, val, src in rates:
    put(ws, row, 1, item); put(ws, row, 2, basis); put(ws, row, 3, val, BLUE, GBP); put(ws, row, 4, src)
widths(ws, [34, 40, 22, 70]); ws.freeze_panes = "A4"

# ---- Terms
ws = wb.create_sheet("Terms")
put(ws, 1, 1, "Invoicing terms — inputs (blue). Add a row per new term; the register for each term goes in Invoices.", TITLE)
header(ws, 3, ["Term ID", "Description", "Saturday sessions", "Wednesday sessions", "Invoice series", "Register loaded?", "Receipts window from", "Receipts window to", "Note"])
import datetime as _dt
terms = [
 (4, TERM_PRIOR, "Autumn 2025 (Sep–Dec)", 11, 11, "2025-5xx / 2025-6xx", "No", _dt.date(2025, 8, 1), _dt.date(2025, 12, 14),
  "Sessions inferred from receipts (£214.50 = 11 × £19.50). Invoice register not supplied — add it to allocate autumn receipts."),
 (5, TERM_CUR, "Term starting January 2026", 21, 20, "2026-xxx", "Yes", _dt.date(2025, 12, 15), _dt.date(2026, 8, 31),
  "21 = £409.50 ÷ £19.50. Wednesday 20 = £320 ÷ £16 (single club invoice). Confirm both with the office."),
]
for row, *vals in terms:
    for j, v in enumerate(vals, 1):
        put(ws, row, j, v, BLUE if j in (3, 4, 5, 6, 7, 8) else BLACK, "dd/mm/yyyy" if j in (7, 8) else None)
widths(ws, [12, 28, 12, 12, 20, 12, 14, 14, 90]); ws.freeze_panes = "A4"

# ---- Families
ws = wb.create_sheet("Families")
put(ws, 1, 1, "Families — stable key. One family can have several invoice numbers over time; the family ID never changes.", TITLE)
header(ws, 3, ["Family ID", "Family surname(s)", "Children on roster", "Children billed (JAN-2026)", "Invoice no(s) JAN-2026", "Bank payer alias (fill in; separate with ;)", "Payer names seen with a verified reference (engine)", "Note"])
for i, r in fam.iterrows():
    row = 4 + i
    put(ws, row, 1, r.fid, BLUE); put(ws, row, 2, r["name"], BLUE)
    put(ws, row, 3, f'=COUNTIF(Children!$B$4:$B$500,$A{row})')
    put(ws, row, 4, f'=COUNTIFS(Children!$B$4:$B$500,$A{row},Children!$F$4:$F$500,"Enrolled")')
    put(ws, row, 5, r.invs, BLUE); put(ws, row, 6, "", fill=Y_FILL); put(ws, row, 7, ""); put(ws, row, 8, r.note, BLUE)
widths(ws, [11, 32, 10, 12, 22, 30, 40, 60]); ws.freeze_panes = "A4"; ws.auto_filter.ref = f"A3:H{3+len(fam)}"

# ---- Children
ws = wb.create_sheet("Children")
put(ws, 1, 1, "Children — facts per child (blue) drive the expected fee (black). 'Check' must be OK; anything else is an invoice that does not follow the rules.", TITLE)
cols = ["Child ID", "Family ID", "Surname", "First name", "Class", "Status", "Term", "Invoice no (roster)",
        "Sessions", "Sessions override", "Siblings billed", "Rate per session", "Tuition", "Discount", "Add-ons charged",
        "Reg fee", "Supplies", "Expected total", "Invoiced on roster", "Difference", "Check", "Note"]
header(ws, 3, cols)
first, last = 4, 3 + len(ch)
for i, r in ch.iterrows():
    row = first + i
    for j, v in enumerate([r.cid, r.fid, r.sur, r.first, r.cls, r.status, TERM_CUR, r.inv_raw], 1):
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
    if r.status == "Enrolled":
        exp_rule = round(r.r_fees - r.discount + (25.0 if r.reg_flag == "Y" else 0) + (25.0 if r.sup_flag == "Y" else 0), 2)
        roster_tot = round(r.r_fees + r.r_sup + r.r_reg - r.r_disc, 2)
        if abs(exp_rule - roster_tot) >= 0.02:
            for c in range(1, 23): ws.cell(row, c).fill = PatternFill("solid", fgColor="FCE4E4")
widths(ws, [9, 10, 20, 14, 10, 13, 10, 12, 8, 9, 8, 10, 10, 9, 14, 8, 9, 11, 11, 10, 8, 55])
ws.freeze_panes = "E4"; ws.auto_filter.ref = f"A3:V{last}"
ws["J3"].comment = Comment("Leave blank for a standard term. Enter sessions only when the child is billed for a different number, e.g. 9 sessions, or 10.5 for half a sibling share.", "ledger")
ws.conditional_formatting.add(f"A4:V{last}", FormulaRule(formula=[f'$U4="CHECK"'], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=F, size=10, color="9C0006")))

# ---- Receipts
ws = wb.create_sheet("Receipts")
put(ws, 1, 1, "Receipts — one row per bank line. Blue = from the bank file. Yellow = your decision. Everything else is written by the engine.", TITLE)
header(ws, 3, ["Date", "Amount", "Payer (bank)", "Reference (bank)", "Term", "Family override (fill in)", "Family",
               "Allocated", "Tier", "Candidates", "Why", "Amount check", "Flags", "Full memo"])
for i, r in rc.iterrows():
    row = 4 + i
    put(ws, row, 1, r.date, BLUE, "dd/mm/yyyy"); put(ws, row, 2, r.amount, BLUE, GBP)
    put(ws, row, 3, r.payer, BLUE); put(ws, row, 4, r.ref, BLUE); put(ws, row, 5, r.term, BLUE)
    put(ws, row, 6, "", fill=Y_FILL)
    put(ws, row, 7, f'=IF($F{row}<>"",$F{row},$J{row})')
    put(ws, row, 8, f'=IF($G{row}<>"","yes","")')
    put(ws, row, 13, r.flag or r.extra, BLUE); put(ws, row, 14, r.memo, BLUE)
widths(ws, [11, 11, 22, 24, 10, 14, 11, 9, 6, 16, 60, 20, 40, 46])
ws.freeze_panes = "C4"; ws.auto_filter.ref = f"A3:N{3 + len(rc)}"

wb.save(OUT)
print("saved", OUT, "| children", len(ch), "| families", len(fam), "| receipts", len(rc))
print("statuses", ch.status.value_counts().to_dict())
print("shared invoices", shared_inv, "| no-invoice children", n_noinv, "| discounts", n_disc, "| out of seq", n_out_of_seq)
print("receipt auto-match", (rc.auto != "").sum(), "| suggested", (rc.suggest != "").sum(), "| basis counts", rc.how.value_counts().to_dict())
