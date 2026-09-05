"""
reconcile.py - allocate bank receipts to families and write the Review, Position and Summary sheets.

    python reconcile.py [fee_ledger.xlsx]

The matching logic lives in engine.py and has no Excel in it; this file is the I/O around it.
Re-run after any change to the Receipts sheet (new bank lines, new overrides). Safe to run repeatedly.
"""
import sys
import datetime as dt
from collections import Counter, defaultdict

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter

from engine import Config, load_roster, read_receipts, allocate, amount_pattern

# ------------------------------------------------------------------ workbook styling
F = "Arial"
BLUE, BLACK, GREEN = Font(name=F, size=10, color="0000FF"), Font(name=F, size=10), Font(name=F, size=10, color="008000")
BOLD, TITLE = Font(name=F, size=10, bold=True), Font(name=F, size=13, bold=True)
H_FILL, Y_FILL = PatternFill("solid", fgColor="D9D9D9"), PatternFill("solid", fgColor="FFFF00")
GREEN_L, AMBER_L, AMBER, RED_L = (PatternFill("solid", fgColor="E2EFDA"), PatternFill("solid", fgColor="FFF2CC"),
                                  PatternFill("solid", fgColor="FFEB9C"), PatternFill("solid", fgColor="FCE4E4"))
GBP = '£#,##0.00;(£#,##0.00);-'


def put(ws, r, c, v, font=BLACK, fmt=None, fill=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


def header(ws, row, cols):
    for j, h in enumerate(cols, 1):
        c = put(ws, row, j, h, BOLD, fill=H_FILL)
        c.alignment = Alignment(wrap_text=True, vertical="center")


def widths(ws, w):
    for j, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(j)].width = x


def write_outputs(wb, rows, roster, cfg, report_date):
    """Write the engine's decisions into Receipts and build the Review, Position and Summary sheets."""
    ws = wb["Receipts"]
    AUTO_TIERS, TERM_CUR, TERM_PRIOR, REPORT_DATE = cfg.auto_tiers, cfg.term_cur, cfg.term_prior, report_date
    fam_name, expected, billed = roster.fam_name, roster.expected, roster.billed

    # --------------------------------------------- write engine columns (I..M: tier..amount check)
    # Flags (N) and Full memo (O) belong to build_ledger.py and are left alone. Writing the amount
    # check into N used to overwrite the bank diagnostics it had put there.
    for x in rows:
        r = x.r
        put(ws, r, 9, x.tier, GREEN)
        put(ws, r, 10, x.fid if x.tier in AUTO_TIERS else "", GREEN)
        put(ws, r, 11, ", ".join(sorted(x.cands)) if x.cands else "", GREEN)
        put(ws, r, 12, x.why, GREEN)
        put(ws, r, 13, "; ".join(y for y in [x.amt, amount_pattern(x.amount, cfg)] if y and y != "n/a"), GREEN)
    for x in rows:
        fill = AMBER if x.tier == "X" else (GREEN_L if x.tier in AUTO_TIERS else AMBER_L)
        for c in range(1, 16): ws.cell(x.r, c).fill = fill
    last_rc = rows[-1].r
    ws.conditional_formatting.add(f"A4:N{last_rc}", FormulaRule(formula=['$I4="X"'], fill=AMBER, font=Font(name=F, size=10, color="9C5700")))
    ws.conditional_formatting.add(f"A4:N{last_rc}", FormulaRule(formula=['AND($H4="",$I4<>"X")'], fill=AMBER_L))
    ws.conditional_formatting.add(f"A4:N{last_rc}", FormulaRule(formula=['$H4="yes"'], fill=GREEN_L))
    # aliases seen -> Families column G
    seen = defaultdict(set)
    for x in rows:
        if x.tier in ("A", "M") and x.fid: seen[x.fid].add(x.payer.strip())
    wf = wb["Families"]
    for r in range(4, wf.max_row + 1):
        fid = wf.cell(r, 1).value
        if fid: put(wf, r, 7, "; ".join(sorted(seen.get(fid, []))), GREEN)

    # ------------------------------------------------------------------ Review sheet
    for name in ("Review", "Arrears", "Position", "Summary"):
        if name in wb.sheetnames: del wb[name]
    wr = wb.create_sheet("Review")
    put(wr, 1, 1, "Review queue  -  receipts the engine would not allocate. Decide, then type the family ID into Receipts column F (row shown) and re-run.", TITLE)
    header(wr, 3, ["Row in Receipts", "Date", "Amount", "Payer (bank)", "Reference (bank)", "Term", "Tier", "Most likely family", "Why the engine stopped", "Amount check", "That family's expected fee"])
    queue = sorted([x for x in rows if x.tier not in AUTO_TIERS], key=lambda x: (-abs(x.amount)))
    for i, x in enumerate(queue):
        r = 4 + i
        single = next(iter(x.cands)) if len(x.cands) == 1 else ""
        vals = [x.r, x.date, x.amount, x.payer, x.ref, x.term, x.tier,
                ", ".join(f"{c} {fam_name.get(c, '')}" for c in sorted(x.cands)), x.why, x.amt,
                expected.get(single, "") if single else ""]
        for j, v in enumerate(vals, 1):
            put(wr, r, j, v, BLUE, "dd/mm/yyyy" if j == 2 else GBP if j in (3, 11) else None)
    for i, x in enumerate(queue):
        fill = AMBER if x.tier == "X" else (AMBER_L if x.tier == "C" else RED_L)
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
        if x.tier in AUTO_TIERS and x.fid and x.term == TERM_CUR: recs[x.fid].append(x)
    prior_money = defaultdict(float)
    for x in rows:
        if x.tier in AUTO_TIERS and x.fid and x.term == TERM_PRIOR:
            if (x.date and x.date.year >= int(cfg.year_cur)) or amount_pattern(x.amount, cfg) == f"{TERM_CUR} amount":
                prior_money[x.fid] += x.amount
    in_review = Counter()
    for x in rows:
        if x.tier not in AUTO_TIERS:
            for c in x.cands: in_review[c] += 1
    order = sorted([f for f in expected if expected[f] > 0], key=lambda f: (-(expected[f] - sum(y.amount for y in recs[f])), fam_name[f]))
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
        _dates = [y.date for y in recs[fid] if y.date]
        put(wp, r, 10, max(_dates) if _dates else "", GREEN, "dd/mm/yyyy")
        put(wp, r, 11, f'=IF($G{r}>0.5,$B$2-Terms!$G$5,"")', fmt="0")
        put(wp, r, 12, " | ".join(sorted({(y.ref or "no reference").strip() for y in recs[fid]}))[:200], GREEN)
        look = []
        if in_review[fid]: look.append(f"{in_review[fid]} receipt(s) in Review name this family")
        if prior_money[fid]: look.append(f"£{prior_money[fid]:,.2f} tagged {TERM_PRIOR} looks like a {TERM_CUR} payment")
        put(wp, r, 13, "; ".join(look), GREEN)
    # static fills (conditional formatting is often dropped on import by other spreadsheet apps)
    for i, fid in enumerate(order):
        r = 5 + i
        got = sum(y.amount for y in recs[fid]); bal = expected[fid] - got
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
    wb.move_sheet("Position", offset=-(len(wb.sheetnames) - 2))
    return queue, order


def main(argv):
    ledger = argv[1] if len(argv) > 1 else "examples/fee_ledger.xlsx"
    cfg = Config()
    wb = load_workbook(ledger)
    roster = load_roster(wb, cfg)
    rows = allocate(read_receipts(wb["Receipts"], cfg), roster, cfg)
    queue, order = write_outputs(wb, rows, roster, cfg, dt.date.today())
    wb.save(ledger)

    tiers = Counter(x.tier for x in rows)
    auto = sum(1 for x in rows if x.tier in cfg.auto_tiers)
    print(f"tiers {dict(sorted(tiers.items()))} | allocated {auto}/{len(rows)} = {auto/len(rows):.0%} | review {len(queue)} | arrears rows {len(order)}")
    got = sum(x.amount for x in rows if x.tier in cfg.auto_tiers and x.term == cfg.term_cur)
    due = sum(x.amount for x in rows if x.term == cfg.term_cur and x.amount > 0)
    print(f"{cfg.term_cur}: allocated £{got:,.2f} of £{due:,.2f}")


if __name__ == "__main__":
    main(sys.argv)
