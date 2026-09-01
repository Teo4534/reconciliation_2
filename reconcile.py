"""
reconcile.py  -  allocation engine and arrears for fee_ledger_2026.xlsx

Reads the ledger built by build_ledger.py, decides which family each bank receipt belongs to,
writes the decision (with a confidence tier and the reasoning) into the Receipts sheet, and adds
three sheets: Review (what a human must decide), Arrears (who owes what, with warnings) and Summary.

Tiers  -  highest evidence first, first hit wins:
  M  manual override typed into Receipts column M          -> allocated
  A  current-term invoice reference naming exactly one family, not contradicted by a surname
     in the memo; OR a bank payer name already seen with a verified reference (alias)   -> allocated
  B  exactly one roster surname found in the memo                                       -> allocated
  C  fuzzy evidence: truncated / misspelt surname, or a child's first name              -> Review
  D  nothing usable, or several families fit                                            -> Review
  X  the reference and the memo point to different families                             -> Review
Change AUTO_TIERS below if you want B reviewed too.

Re-run after any change to Receipts (new bank lines, new overrides). Safe to run repeatedly.
"""
import re, sys, unicodedata, datetime as dt
from collections import defaultdict, Counter
from difflib import SequenceMatcher
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter

LEDGER = sys.argv[1] if len(sys.argv) > 1 else "examples/fee_ledger.xlsx"
AUTO_TIERS = {"M", "A", "B"}
TERM_CUR, TERM_PRIOR = "JAN-2026", "AUT-2025"
REPORT_DATE = dt.date.today()
STOP = {"THE", "AND", "FOR", "INV", "FEES", "FEE", "PART", "PAYM", "PAYMENT", "SCHOOL", "ECOLE", "TWINS", "MRS",
        "MISS", "LTD", "BGC", "BBP", "REG", "TERM", "FRENCH", "SAT", "SATURDAY", "CLUB", "AND", "SON", "VAN", "DER"}

def norm(x):
    return unicodedata.normalize("NFKD", str(x or "")).encode("ascii", "ignore").decode().upper()
def tokens(text):
    return [t for t in re.findall(r"[A-Z]{3,}", norm(text)) if t not in STOP]
def payer_key(p):
    return re.sub(r"[^A-Z]", "", norm(p))

wb = load_workbook(LEDGER)
F = "Arial"
BLUE, BLACK, GREEN = Font(name=F, size=10, color="0000FF"), Font(name=F, size=10), Font(name=F, size=10, color="008000")
BOLD, TITLE = Font(name=F, size=10, bold=True), Font(name=F, size=13, bold=True)
H_FILL, Y_FILL = PatternFill("solid", fgColor="D9D9D9"), PatternFill("solid", fgColor="FFFF00")
GREEN_L, AMBER_L, AMBER, RED_L = (PatternFill("solid", fgColor="E2EFDA"), PatternFill("solid", fgColor="FFF2CC"),
                                  PatternFill("solid", fgColor="FFEB9C"), PatternFill("solid", fgColor="FCE4E4"))
GBP = '£#,##0.00;(£#,##0.00);-'

def put(ws, r, c, v, font=BLACK, fmt=None, fill=None):
    cell = ws.cell(row=r, column=c, value=v); cell.font = font
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    return cell
def header(ws, row, cols):
    for j, h in enumerate(cols, 1):
        c = put(ws, row, j, h, BOLD, fill=H_FILL); c.alignment = Alignment(wrap_text=True, vertical="center")
def widths(ws, w):
    for j, x in enumerate(w, 1): ws.column_dimensions[get_column_letter(j)].width = x

# ------------------------------------------------------------------ roster indexes
fam_name, fam_alias_user = {}, defaultdict(set)
for row in wb["Families"].iter_rows(min_row=4, values_only=True):
    if not row[0]: continue
    fam_name[row[0]] = row[1]
    for a in str(row[5] or "").split(";"):
        if payer_key(a): fam_alias_user[payer_key(a)].add(row[0])

sur_index, first_count, first_fid, inv_index = defaultdict(set), Counter(), {}, defaultdict(set)
expected, billed = defaultdict(float), Counter()
for row in wb["Children"].iter_rows(min_row=4, values_only=True):
    if not row[0]: continue
    cid, fid, sur, first, cls, status, term, inv = row[:8]
    invoiced = row[18]
    invoiced = float(invoiced) if isinstance(invoiced, (int, float)) else 0.0
    for part in re.split(r"[-\s]", norm(sur)):
        if len(part) >= 3: sur_index[part].add(fid)
    for fn in re.split(r"[-\s]", norm(first)):
        if len(fn) >= 4: first_count[fn] += 1; first_fid[fn] = fid
    if inv: inv_index[re.sub(r"[a.]$", "", str(inv).strip())].add(fid)
    if status == "Enrolled" and term == TERM_CUR:
        expected[fid] += invoiced; billed[fid] += 1
unique_first = {fn: fid for fn, fid in first_fid.items() if first_count[fn] == 1}
sur_parts = [p for p in sur_index if len(p) >= 5]

def evidence(text):
    """fid -> list of signals found in text: exact surname parts, unique child first names, fuzzy surname hits"""
    ev = defaultdict(list)
    for t in tokens(text):
        if t in sur_index:
            for f in sur_index[t]: ev[f].append(f"surname {t}")
            continue
        if t in unique_first: ev[unique_first[t]].append(f"child first name {t.title()}")
        for p_ in sur_parts:
            if p_.startswith(t) and len(t) >= 4:
                for f in sur_index[p_]: ev[f].append(f"'{t}' starts {p_}")
            elif len(t) >= 5 and SequenceMatcher(None, t, p_).ratio() >= 0.85:
                for f in sur_index[p_]: ev[f].append(f"'{t}' ~ {p_}")
            elif len(t) >= 8 and p_ in t:
                for f in sur_index[p_]: ev[f].append(f"{p_} inside '{t}'")
    return ev
def families_named(text):
    return {f for f, sig in evidence(text).items() if any(s.startswith("surname ") for s in sig)}
def decide(text, source=""):
    """returns (tier, fid, candidates, reason) from name evidence alone"""
    ev = evidence(text)
    if not ev: return "D", "", set(), "no roster surname or child name" + source
    exact = {f for f, sig in ev.items() if any(s.startswith("surname ") for s in sig)}
    pool = exact if exact else set(ev)
    score = {f: len(set(ev[f])) for f in pool}
    top = max(score.values()); winners = [f for f in pool if score[f] == top]
    if len(winners) > 1:                                   # tie-break: prefer families actually billed this term
        w2 = [f for f in winners if billed.get(f, 0) > 0]
        if len(w2) == 1: winners, tie = w2, True
        else: return "D", "", pool, ("several families fit: " + "; ".join(f"{f} ({', '.join(sorted(set(ev[f])))})" for f in sorted(pool)) + source)
    else: tie = False
    f = winners[0]; why = ", ".join(sorted(set(ev[f]))) + source
    if tie: return "C", "", {f}, why + "  -  chosen over " + ", ".join(sorted(pool - {f})) + " because only this family is billed this term"
    if f in exact and (len(pool) == 1 or top >= 2): return "B", f, set(), why
    return "C", "", {f}, why + ("  -  other fits: " + ", ".join(sorted(pool - {f})) if len(pool) > 1 else "")
AUT_BASE, JAN_BASE = [214.5, 379.5, 506.0], [409.5, 724.5, 966.0]     # 11 and 21 sessions x sibling tiers
def amount_pattern(amount):
    for base, label in ((AUT_BASE, f"{TERM_PRIOR} amount"), (JAN_BASE, f"{TERM_CUR} amount")):
        for b in base:
            for addon in (0, 25, 50, 75):
                if abs(amount - (b + addon)) < 0.51 or abs(amount - b / 2) < 0.51: return label
    return ""
def amount_check(amount, fid, term):
    if term != TERM_CUR or fid not in expected: return "n/a"
    e = expected[fid]; n = billed[fid] or 1
    plausible = [(e, "matches expected"), (e / 2, "half of expected"), (e / n, "one child's share"), (e - 25, "expected less £25"),
                 (e - 50, "expected less £50"), (e + 25, "expected plus £25"), (e + 50, "expected plus £50")]
    for v, why in plausible:
        if abs(amount - v) < 0.51: return why
    return f"unusual (expected £{e:,.2f})"

# ------------------------------------------------------------------ receipts, pass 1: references and aliases
ws = wb["Receipts"]
rows = []
for r in range(4, ws.max_row + 1):
    if ws.cell(r, 1).value is None: break
    g = lambda c: ws.cell(r, c).value
    memo = g(14) or ""
    M = norm(memo)
    m = re.search(r"2026\s*-?\s*(\d{3})", M) or re.search(r"(?<!\d)26(\d{3})(?!\d)", M)
    cur_ref = f"2026-{m.group(1)}" if m else ""
    pm = re.search(r"2025\s*-?\s*\d{3}|(?<!\d)2[1-4]\s*-?\s*1\d\d(?!\d)|(?<!\d)2[1-4]\d{3}(?!\d)", M)
    prior_ref = pm.group(0) if (pm and not cur_ref) else ""
    rows.append(dict(r=r, date=g(1), amount=float(g(2) or 0), payer=g(3) or "", ref=g(4) or "", memo=memo,
                     cur_ref=cur_ref, prior_ref=prior_ref, term=g(5), override=str(g(6) or "").strip(), note=g(13) or ""))

alias = defaultdict(set)                       # payer key -> families, learned from verified references / overrides
for k, v in fam_alias_user.items(): alias[k] |= v
for x in rows:
    memo_fams = families_named(x["memo"])
    x["named"] = memo_fams
    x["tier"] = x["fid"] = x["why"] = ""; x["cands"] = set()
    if x["override"]:
        x["tier"], x["fid"], x["why"] = "M", x["override"], "manual override"
        alias[payer_key(x["payer"])].add(x["override"]); continue
    if x["cur_ref"]:
        f = inv_index.get(x["cur_ref"], set())
        if len(f) > 1 and len(f & memo_fams) == 1: f = f & memo_fams
        if len(f) == 1:
            cand = next(iter(f))
            if memo_fams and cand not in memo_fams:
                x["tier"], x["cands"], x["why"] = "X", memo_fams | {cand}, (
                    f"reference {x['cur_ref']} points to {cand} but the memo names " + ", ".join(sorted(memo_fams)) +
                    "  -  parent may have typed the wrong invoice number")
            else:
                x["tier"], x["fid"], x["why"] = "A", cand, f"invoice reference {x['cur_ref']}"
                alias[payer_key(x["payer"])].add(cand)
        elif len(f) > 1:
            x["tier"], x["cands"], x["why"] = "D", f, f"reference {x['cur_ref']} is shared by " + ", ".join(sorted(f))
        else:
            x["why"] = f"reference {x['cur_ref']} not in roster"
alias = {k: v for k, v in alias.items() if len(k) >= 4 and len(v) == 1}   # only unambiguous aliases

# ------------------------------------------------------------------ pass 2: names for everything still open
for x in rows:
    if x["tier"]: continue
    pk = payer_key(x["payer"]); named = x["named"]
    if pk in alias:
        x["tier"], x["fid"] = "A", next(iter(alias[pk]))
        x["why"] = (x["why"] + "; " if x["why"] else "") + "payer name previously seen with a verified reference for this family"
    else:
        tier, fid, cands, why = decide(x["memo"])
        if tier == "D" and not cands and x["note"]:
            tier, fid, cands, why = decide(x["note"], " (from the bank-sheet note, not the memo)")
        x["tier"], x["fid"], x["cands"] = tier, fid, cands
        x["why"] = (x["why"] + "; " if x["why"] else "") + why
    if x["prior_ref"] and x["tier"] in ("A", "B", "C") :
        x["why"] += f" (prior-term reference {x['prior_ref']}  -  family identified, invoice not loaded)"
for x in rows:
    x["amt"] = amount_check(x["amount"], x["fid"] or (next(iter(x["cands"])) if len(x["cands"]) == 1 else ""), x["term"])

# ------------------------------------------------------------------ write engine columns (P..T)
for x in rows:
    r = x["r"]
    put(ws, r, 9, x["tier"], GREEN)
    put(ws, r, 10, x["fid"] if x["tier"] in AUTO_TIERS else "", GREEN)
    put(ws, r, 11, ", ".join(sorted(x["cands"])) if x["cands"] else "", GREEN)
    put(ws, r, 12, x["why"], GREEN); put(ws, r, 13, ("; ".join(y for y in [x["amt"], amount_pattern(x["amount"])] if y and y != "n/a")) or (ws.cell(r, 13).value or ""), GREEN)
for x in rows:
    fill = AMBER if x["tier"] == "X" else (GREEN_L if x["tier"] in AUTO_TIERS else AMBER_L)
    for c in range(1, 15): ws.cell(x["r"], c).fill = fill
last_rc = rows[-1]["r"]
ws.conditional_formatting.add(f"A4:N{last_rc}", FormulaRule(formula=['$I4="X"'], fill=AMBER, font=Font(name=F, size=10, color="9C5700")))
ws.conditional_formatting.add(f"A4:N{last_rc}", FormulaRule(formula=['AND($H4="",$I4<>"X")'], fill=AMBER_L))
ws.conditional_formatting.add(f"A4:N{last_rc}", FormulaRule(formula=['$H4="yes"'], fill=GREEN_L))
# aliases seen -> Families column G
seen = defaultdict(set)
for x in rows:
    if x["tier"] in ("A", "M") and x["fid"]: seen[x["fid"]].add(x["payer"].strip())
wf = wb["Families"]
for r in range(4, wf.max_row + 1):
    fid = wf.cell(r, 1).value
    if fid: put(wf, r, 7, "; ".join(sorted(seen.get(fid, []))), GREEN)

# ------------------------------------------------------------------ Review sheet
for name in ("Review", "Arrears", "Position", "Summary"):
    if name in wb.sheetnames: del wb[name]
wr = wb.create_sheet("Review")
put(wr, 1, 1, "Review queue  -  receipts the engine would not allocate. Decide, then type the family ID into Receipts column M (row shown) and re-run.", TITLE)
header(wr, 3, ["Row in Receipts", "Date", "Amount", "Payer (bank)", "Reference (bank)", "Term", "Tier", "Most likely family", "Why the engine stopped", "Amount check", "That family's expected fee"])
queue = sorted([x for x in rows if x["tier"] not in AUTO_TIERS], key=lambda x: (-abs(x["amount"])))
for i, x in enumerate(queue):
    r = 4 + i
    single = next(iter(x["cands"])) if len(x["cands"]) == 1 else ""
    vals = [x["r"], x["date"], x["amount"], x["payer"], x["ref"], x["term"], x["tier"],
            ", ".join(f"{c} {fam_name.get(c, '')}" for c in sorted(x["cands"])), x["why"], x["amt"],
            expected.get(single, "") if single else ""]
    for j, v in enumerate(vals, 1):
        put(wr, r, j, v, BLUE, "dd/mm/yyyy" if j == 2 else GBP if j in (3, 11) else None)
for i, x in enumerate(queue):
    fill = AMBER if x["tier"] == "X" else (AMBER_L if x["tier"] == "C" else RED_L)
    for c in range(1, 12): wr.cell(4 + i, c).fill = fill
widths(wr, [9, 11, 11, 22, 24, 10, 6, 34, 70, 22, 14]); wr.freeze_panes = "D4"; wr.auto_filter.ref = f"A3:K{3 + len(queue)}"
lastq = 3 + len(queue)
wr.conditional_formatting.add(f"A4:K{lastq}", FormulaRule(formula=['$G4="X"'], fill=AMBER, font=Font(name=F, size=10, bold=True, color="9C5700")))
wr.conditional_formatting.add(f"A4:K{lastq}", FormulaRule(formula=['$G4="C"'], fill=AMBER_L))
wr.conditional_formatting.add(f"A4:K{lastq}", FormulaRule(formula=['$G4="D"'], fill=RED_L))

# ------------------------------------------------------------------ Position sheet (the answer sheet)
wp = wb.create_sheet("Position")
put(wp, 1, 1, f"Position  -  {TERM_CUR}. One row per family. Green = settled. Red = owing. Amber = check the last two columns before chasing.", TITLE)
put(wp, 2, 1, "Report date"); put(wp, 2, 2, REPORT_DATE, BLUE, "dd/mm/yyyy")
header(wp, 4, ["Family ID", "Family", "Children", "Invoice no(s)", "Expected", "Received", "Balance", "Status",
               "Receipts", "Last payment", "Days", "References used (bank)", "Needs a look"])
recs = defaultdict(list)
for x in rows:
    if x["tier"] in AUTO_TIERS and x["fid"] and x["term"] == TERM_CUR: recs[x["fid"]].append(x)
prior_money = defaultdict(float)
for x in rows:
    if x["tier"] in AUTO_TIERS and x["fid"] and x["term"] == TERM_PRIOR:
        if (x["date"] and x["date"].year >= 2026) or amount_pattern(x["amount"]) == f"{TERM_CUR} amount":
            prior_money[x["fid"]] += x["amount"]
in_review = Counter()
for x in rows:
    if x["tier"] not in AUTO_TIERS:
        for c in x["cands"]: in_review[c] += 1
order = sorted([f for f in expected if expected[f] > 0], key=lambda f: (-(expected[f] - sum(y["amount"] for y in recs[f])), fam_name[f]))
inv_no = {}
for row in wb["Families"].iter_rows(min_row=4, values_only=True):
    if row[0]: inv_no[row[0]] = row[4]
for i, fid in enumerate(order):
    r = 5 + i
    put(wp, r, 1, fid, BLUE); put(wp, r, 2, fam_name[fid], BLUE); put(wp, r, 3, billed[fid], BLUE)
    put(wp, r, 4, inv_no.get(fid, ""), BLUE)
    put(wp, r, 5, f'=SUMIFS(Children!$R$4:$R$500,Children!$B$4:$B$500,$A{r},Children!$F$4:$F$500,"Enrolled",Children!$G$4:$G$500,"{TERM_CUR}")', fmt=GBP)
    put(wp, r, 6, f'=SUMIFS(Receipts!$B$4:$B$600,Receipts!$G$4:$G$600,$A{r},Receipts!$E$4:$E$600,"{TERM_CUR}")', fmt=GBP)
    put(wp, r, 7, f'=$E{r}-$F{r}', fmt=GBP)
    put(wp, r, 8, f'=IF($E{r}=0,"no charge",IF($F{r}<=0,"unpaid",IF(ABS($G{r})<=0.5,"settled",IF($G{r}>0,"part paid","overpaid"))))')
    put(wp, r, 9, f'=COUNTIFS(Receipts!$G$4:$G$600,$A{r},Receipts!$E$4:$E$600,"{TERM_CUR}")', fmt="0")
    _dates = [y["date"] for y in recs[fid] if y["date"]]
    put(wp, r, 10, max(_dates) if _dates else "", GREEN, "dd/mm/yyyy")
    put(wp, r, 11, f'=IF($G{r}>0.5,$B$2-Terms!$G$5,"")', fmt="0")
    put(wp, r, 12, " | ".join(sorted({(y["ref"] or "no reference").strip() for y in recs[fid]}))[:200], GREEN)
    look = []
    if in_review[fid]: look.append(f"{in_review[fid]} receipt(s) in Review name this family")
    if prior_money[fid]: look.append(f"£{prior_money[fid]:,.2f} tagged {TERM_PRIOR} looks like a {TERM_CUR} payment")
    put(wp, r, 13, "; ".join(look), GREEN)
# static fills (conditional formatting is often dropped on import by other spreadsheet apps)
for i, fid in enumerate(order):
    r = 5 + i
    got = sum(y["amount"] for y in recs[fid]); bal = expected[fid] - got
    look = wp.cell(r, 13).value
    fill = AMBER_L if look else (GREEN_L if abs(bal) <= 0.5 else RED_L)
    for c in range(1, 14): wp.cell(r, c).fill = fill
last_p = 4 + len(order)
rng = f"A5:M{last_p}"
wp.conditional_formatting.add(rng, FormulaRule(formula=['$M5<>""'], fill=AMBER_L))
wp.conditional_formatting.add(rng, FormulaRule(formula=['OR($H5="settled",$H5="no charge")'], fill=GREEN_L))
wp.conditional_formatting.add(rng, FormulaRule(formula=['$G5>0.5'], fill=RED_L))
wp.conditional_formatting.add(f"K5:K{last_p}", CellIsRule(operator="greaterThan", formula=["90"], font=Font(name=F, size=10, bold=True, color="9C0006")))
widths(wp, [10, 28, 8, 14, 11, 11, 11, 10, 8, 12, 6, 34, 48])
wp.freeze_panes = "C5"; wp.auto_filter.ref = f"A4:M{last_p}"

# ------------------------------------------------------------------ Summary
wsum = wb.create_sheet("Summary")
put(wsum, 1, 1, "Summary  -  all formulas", TITLE)
RB, RI = "Receipts!$B$4:$B$600", "Receipts!$E$4:$E$600"
RP, RO = "Receipts!$I$4:$I$600", "Receipts!$H$4:$H$600"
items = [("Receipts", None, None),
         ("Bank lines", f'=COUNT({RB})', "0"), ("Money in", f'=SUMIF({RB},">0")', GBP), ("Money out (refunds / outflows)", f'=SUMIF({RB},"<0")', GBP),
         (f"Money in  -  {TERM_PRIOR}", f'=SUMIFS({RB},{RI},"{TERM_PRIOR}",{RB},">0")', GBP),
         (f"Money in  -  {TERM_CUR}", f'=SUMIFS({RB},{RI},"{TERM_CUR}",{RB},">0")', GBP),
         ("", None, None), ("Allocation by tier (all terms)", None, None)]
for t, label in [("M", "M  -  manual override"), ("A", "A  -  reference or verified payer alias"), ("B", "B  -  unique surname in memo"),
                 ("C", "C  -  fuzzy / first name (review)"), ("D", "D  -  nothing usable or ambiguous (review)"), ("X", "X  -  reference contradicts memo (review)")]:
    items.append((label + "  [lines]", f'=COUNTIF({RP},"{t}")', "0"))
    items.append((label + "  [£ in]", f'=SUMIFS({RB},{RP},"{t}",{RB},">0")', GBP))
items += [("Lines allocated automatically", f'=COUNTIF({RO},"yes")', "0"),
          ("Share of lines allocated automatically", f'=IF(COUNT({RB})=0,0,COUNTIF({RO},"yes")/COUNT({RB}))', "0.0%"),
          (f"{TERM_CUR} money in allocated", f'=SUMIFS({RB},{RI},"{TERM_CUR}",{RO},"yes",{RB},">0")', GBP),
          (f"{TERM_CUR} money in awaiting review", f'=SUMIFS({RB},{RI},"{TERM_CUR}",{RO},"",{RB},">0")', GBP),
          ("Lines in Review queue", '=COUNTA(Review!$A$4:$A$600)', "0"),
          ("", None, None), (f"Position  -  {TERM_CUR}", None, None),
          ("Expected income (rules)", '=SUM(Position!$E$5:$E$600)', GBP), ("Received (allocated)", '=SUM(Position!$F$5:$F$600)', GBP),
          ("Outstanding", '=SUMIF(Position!$G$5:$G$600,">0")', GBP),
          ("Families settled", '=COUNTIF(Position!$H$5:$H$600,"settled")', "0"), ("Families part paid", '=COUNTIF(Position!$H$5:$H$600,"part paid")', "0"),
          ("Families with nothing allocated", '=COUNTIF(Position!$H$5:$H$600,"unpaid")', "0"), ("Families overpaid", '=COUNTIF(Position!$H$5:$H$600,"overpaid")', "0"),
          ("Families flagged 'needs a look'", '=COUNTIF(Position!$M$5:$M$600,"?*")', "0"),
          ("", None, None), ("Rule check", None, None),
          ("Children billed", '=COUNTIF(Children!$F$4:$F$500,"Enrolled")', "0"),
          ("Invoices reproduced exactly by the fee rules", '=COUNTIFS(Children!$F$4:$F$500,"Enrolled",Children!$U$4:$U$500,"OK")', "0"),
          ("", None, None),
          (f"Outstanding is overstated by whatever sits in Review and by any 2026 money tagged to {TERM_PRIOR}. The {TERM_PRIOR} register is not loaded, so autumn has no position yet.", None, None)]
for i, (label, f, fmt) in enumerate(items, 3):
    put(wsum, i, 1, label, BOLD if f is None and label else BLACK)
    if f: put(wsum, i, 2, f, GREEN, fmt)
widths(wsum, [70, 18])

wb.move_sheet("Summary", offset=-(len(wb.sheetnames) - 1))   # Summary first
wb.save(LEDGER)
tiers = Counter(x["tier"] for x in rows)
auto = sum(1 for x in rows if x["tier"] in AUTO_TIERS)
wb.move_sheet("Position", offset=-(len(wb.sheetnames) - 2))
print(f"tiers {dict(sorted(tiers.items()))} | allocated {auto}/{len(rows)} = {auto/len(rows):.0%} | review {len(queue)} | arrears rows {len(order)}")
print(f"{TERM_CUR}: allocated £{sum(x['amount'] for x in rows if x['tier'] in AUTO_TIERS and x['term']==TERM_CUR):,.2f} of £{sum(x['amount'] for x in rows if x['term']==TERM_CUR and x['amount']>0):,.2f}")
