"""
engine.py - the allocation logic, with no Excel writing and no side effects.

Everything here is a plain function of its inputs, so it can be unit-tested in milliseconds
without generating a workbook. reconcile.py loads the ledger, calls into this module, and
writes the result back out.

Tiers, highest evidence first, first hit wins:
    M  manual override typed into the Receipts sheet                      -> allocated
    A  current-term invoice reference naming exactly one family, not
       contradicted by a surname in the memo; or a payer name already
       seen with a verified reference (an alias)                          -> allocated
    B  exactly one roster surname found in the memo                       -> allocated
    C  fuzzy evidence: truncated or misspelt surname, or a child's name   -> review
    D  nothing usable, or several families fit                            -> review
    X  the reference and the memo point to different families            -> review
"""
from __future__ import annotations

import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from difflib import SequenceMatcher

# ----------------------------------------------------------------------------- configuration

STOP_WORDS = frozenset({
    "THE", "AND", "FOR", "INV", "FEES", "FEE", "PART", "PAYM", "PAYMENT", "SCHOOL", "ECOLE",
    "TWINS", "MRS", "MISS", "LTD", "BGC", "BBP", "REG", "TERM", "FRENCH", "SAT", "SATURDAY",
    "CLUB", "SON", "VAN", "DER",
})


@dataclass(frozen=True)
class Config:
    """Everything that varies by term or by school. Nothing below this line is hard-coded."""
    term_cur: str = "JAN-2026"
    term_prior: str = "AUT-2025"
    auto_tiers: frozenset = frozenset({"M", "A", "B"})
    fuzzy_ratio: float = 0.85          # SequenceMatcher threshold for a misspelt surname
    min_alias_len: int = 4             # shorter payer keys are too ambiguous to trust as aliases
    # Fee bases per term: 11 and 21 sessions x sibling tiers. Used only to label amounts.
    fee_bases: tuple = (("AUT-2025", (214.5, 379.5, 506.0)), ("JAN-2026", (409.5, 724.5, 966.0)))

    @property
    def year_cur(self) -> str:
        return self.term_cur[-4:]

    @property
    def year_prior(self) -> str:
        return self.term_prior[-4:]

    def ref_patterns(self):
        """Compiled regexes for a current-term reference (two spellings) and a prior-term one."""
        y, yy, py = self.year_cur, self.year_cur[2:], self.year_prior
        cur = (re.compile(rf"{y}\s*-?\s*(\d{{3}})"), re.compile(rf"(?<!\d){yy}(\d{{3}})(?!\d)"))
        prior = re.compile(rf"{py}\s*-?\s*\d{{3}}|(?<!\d)2[1-4]\s*-?\s*1\d\d(?!\d)|(?<!\d)2[1-4]\d{{3}}(?!\d)")
        return cur, prior


# ----------------------------------------------------------------------------- text helpers

def norm(x) -> str:
    """Upper-case ASCII with accents stripped, so ELODIE and Élodie compare equal."""
    return unicodedata.normalize("NFKD", str(x or "")).encode("ascii", "ignore").decode().upper()


def tokens(text) -> list[str]:
    return [t for t in re.findall(r"[A-Z]{3,}", norm(text)) if t not in STOP_WORDS]


def payer_key(payer) -> str:
    return re.sub(r"[^A-Z]", "", norm(payer))


# ----------------------------------------------------------------------------- roster

@dataclass
class Roster:
    fam_name: dict = field(default_factory=dict)            # FAM-nnn -> display name
    alias_seed: dict = field(default_factory=lambda: defaultdict(set))   # payer key -> {fid}, typed by a person
    sur_index: dict = field(default_factory=lambda: defaultdict(set))    # surname part -> {fid}
    unique_first: dict = field(default_factory=dict)        # child's first name -> fid, only if unique
    inv_index: dict = field(default_factory=lambda: defaultdict(set))    # invoice ref -> {fid}
    expected: dict = field(default_factory=lambda: defaultdict(float))   # fid -> £ invoiced this term
    billed: Counter = field(default_factory=Counter)        # fid -> children billed this term

    @property
    def sur_parts(self) -> list[str]:
        return [p for p in self.sur_index if len(p) >= 5]


def load_roster(wb, cfg: Config) -> Roster:
    """Build the lookup indexes from the Families and Children sheets."""
    r = Roster()
    for row in wb["Families"].iter_rows(min_row=4, values_only=True):
        if not row[0]:
            continue
        r.fam_name[row[0]] = row[1]
        for a in str(row[5] or "").split(";"):
            if payer_key(a):
                r.alias_seed[payer_key(a)].add(row[0])

    first_count, first_fid = Counter(), {}
    for row in wb["Children"].iter_rows(min_row=4, values_only=True):
        if not row[0]:
            continue
        _cid, fid, sur, first, _cls, status, term, inv = row[:8]
        invoiced = row[18]
        invoiced = float(invoiced) if isinstance(invoiced, (int, float)) else 0.0
        for part in re.split(r"[-\s]", norm(sur)):
            if len(part) >= 3:
                r.sur_index[part].add(fid)
        for fn in re.split(r"[-\s]", norm(first)):
            if len(fn) >= 4:
                first_count[fn] += 1
                first_fid[fn] = fid
        if inv:
            r.inv_index[re.sub(r"[a.]$", "", str(inv).strip())].add(fid)
        if status == "Enrolled" and term == cfg.term_cur:
            r.expected[fid] += invoiced
            r.billed[fid] += 1
    r.unique_first = {fn: fid for fn, fid in first_fid.items() if first_count[fn] == 1}
    return r


# ----------------------------------------------------------------------------- evidence

def evidence(text, roster: Roster, cfg: Config) -> dict[str, list[str]]:
    """fid -> the signals found in text: exact surname parts, unique child first names, fuzzy hits."""
    ev = defaultdict(list)
    sur_parts = roster.sur_parts
    for t in tokens(text):
        if t in roster.sur_index:
            for f in roster.sur_index[t]:
                ev[f].append(f"surname {t}")
            continue
        if t in roster.unique_first:
            ev[roster.unique_first[t]].append(f"child first name {t.title()}")
        for p in sur_parts:
            if p.startswith(t) and len(t) >= 4:
                for f in roster.sur_index[p]:
                    ev[f].append(f"'{t}' starts {p}")
            elif len(t) >= 5 and SequenceMatcher(None, t, p).ratio() >= cfg.fuzzy_ratio:
                for f in roster.sur_index[p]:
                    ev[f].append(f"'{t}' ~ {p}")
            elif len(t) >= 8 and p in t:
                for f in roster.sur_index[p]:
                    ev[f].append(f"{p} inside '{t}'")
    return ev


def families_named(ev) -> set[str]:
    """Families backed by an exact surname hit, as opposed to fuzzy evidence."""
    return {f for f, sig in ev.items() if any(s.startswith("surname ") for s in sig)}


def decide(ev, roster: Roster, source: str = ""):
    """Turn an evidence dict into (tier, fid, candidates, reason) using name evidence alone."""
    if not ev:
        return "D", "", set(), "no roster surname or child name" + source
    exact = families_named(ev)
    pool = exact if exact else set(ev)
    score = {f: len(set(ev[f])) for f in pool}
    top = max(score.values())
    winners = [f for f in pool if score[f] == top]
    tie = False
    if len(winners) > 1:
        # tie-break: prefer the family actually billed this term
        billed = [f for f in winners if roster.billed.get(f, 0) > 0]
        if len(billed) == 1:
            winners, tie = billed, True
        else:
            detail = "; ".join(f"{f} ({', '.join(sorted(set(ev[f])))})" for f in sorted(pool))
            return "D", "", pool, "several families fit: " + detail + source
    f = winners[0]
    why = ", ".join(sorted(set(ev[f]))) + source
    if tie:
        others = ", ".join(sorted(pool - {f}))
        return "C", "", {f}, f"{why}  -  chosen over {others} because only this family is billed this term"
    if f in exact and (len(pool) == 1 or top >= 2):
        return "B", f, set(), why
    others = ", ".join(sorted(pool - {f}))
    return "C", "", {f}, why + (f"  -  other fits: {others}" if len(pool) > 1 else "")


# ----------------------------------------------------------------------------- amounts

def amount_pattern(amount: float, cfg: Config) -> str:
    """Label an amount as looking like a known term's fee, or return ''."""
    for term, bases in cfg.fee_bases:
        for b in bases:
            for addon in (0, 25, 50, 75):
                if abs(amount - (b + addon)) < 0.51 or abs(amount - b / 2) < 0.51:
                    return f"{term} amount"
    return ""


def amount_check(amount: float, fid: str, term: str, roster: Roster, cfg: Config) -> str:
    """Is this amount plausible for this family's expected fee? A sanity check, not evidence."""
    if term != cfg.term_cur or fid not in roster.expected:
        return "n/a"
    e = roster.expected[fid]
    n = roster.billed[fid] or 1
    plausible = [
        (e, "matches expected"), (e / 2, "half of expected"), (e / n, "one child's share"),
        (e - 25, "expected less £25"), (e - 50, "expected less £50"),
        (e + 25, "expected plus £25"), (e + 50, "expected plus £50"),
    ]
    for v, why in plausible:
        if abs(amount - v) < 0.51:
            return why
    return f"unusual (expected £{e:,.2f})"


# ----------------------------------------------------------------------------- receipts

@dataclass
class Receipt:
    r: int                      # row in the Receipts sheet
    date: object
    amount: float
    payer: str
    ref: str
    memo: str
    cur_ref: str
    prior_ref: str
    term: str
    override: str
    note: str
    named: set = field(default_factory=set)     # families named by an exact surname in the memo
    tier: str = ""
    fid: str = ""
    why: str = ""
    cands: set = field(default_factory=set)
    amt: str = ""


def parse_refs(memo: str, cfg: Config) -> tuple[str, str]:
    """Pull a current-term reference (normalised to YYYY-nnn) and, failing that, a prior-term one."""
    (cur_a, cur_b), prior = cfg.ref_patterns()
    m = cur_a.search(memo) or cur_b.search(memo)
    cur_ref = f"{cfg.year_cur}-{m.group(1)}" if m else ""
    pm = prior.search(memo)
    prior_ref = pm.group(0) if (pm and not cur_ref) else ""
    return cur_ref, prior_ref


def read_receipts(ws, cfg: Config) -> list[Receipt]:
    rows = []
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 1).value is None:
            break
        g = lambda c: ws.cell(r, c).value
        memo = g(14) or ""
        cur_ref, prior_ref = parse_refs(norm(memo), cfg)
        rows.append(Receipt(
            r=r, date=g(1), amount=float(g(2) or 0), payer=g(3) or "", ref=g(4) or "", memo=memo,
            cur_ref=cur_ref, prior_ref=prior_ref, term=g(5), override=str(g(6) or "").strip(),
            note=g(13) or "",
        ))
    return rows


# ----------------------------------------------------------------------------- allocation

def allocate(rows: list[Receipt], roster: Roster, cfg: Config) -> list[Receipt]:
    """Assign a tier, family and reason to every receipt. Two passes:

    1. references and overrides, which also teach us payer aliases
    2. names, for everything still open, using the aliases learned in pass 1
    """
    alias = defaultdict(set)
    for k, v in roster.alias_seed.items():
        alias[k] |= v

    for x in rows:
        x.named = families_named(evidence(x.memo, roster, cfg))
        x.tier = x.fid = x.why = ""
        x.cands = set()
        if x.override:
            x.tier, x.fid, x.why = "M", x.override, "manual override"
            alias[payer_key(x.payer)].add(x.override)
            continue
        if not x.cur_ref:
            continue
        fams = roster.inv_index.get(x.cur_ref, set())
        if len(fams) > 1 and len(fams & x.named) == 1:
            fams = fams & x.named
        if len(fams) == 1:
            cand = next(iter(fams))
            if x.named and cand not in x.named:
                x.tier, x.cands = "X", x.named | {cand}
                x.why = (f"reference {x.cur_ref} points to {cand} but the memo names "
                         + ", ".join(sorted(x.named)) + "  -  parent may have typed the wrong invoice number")
            else:
                x.tier, x.fid, x.why = "A", cand, f"invoice reference {x.cur_ref}"
                alias[payer_key(x.payer)].add(cand)
        elif len(fams) > 1:
            x.tier, x.cands, x.why = "D", fams, f"reference {x.cur_ref} is shared by " + ", ".join(sorted(fams))
        else:
            x.why = f"reference {x.cur_ref} not in roster"

    # only keep aliases that are long enough to be distinctive and point at exactly one family
    alias = {k: v for k, v in alias.items() if len(k) >= cfg.min_alias_len and len(v) == 1}

    for x in rows:
        if x.tier:
            continue
        pk = payer_key(x.payer)
        if pk in alias:
            x.tier, x.fid = "A", next(iter(alias[pk]))
            x.why = _join(x.why, "payer name previously seen with a verified reference for this family")
        else:
            tier, fid, cands, why = decide(evidence(x.memo, roster, cfg), roster)
            if tier == "D" and not cands and x.note:
                tier, fid, cands, why = decide(evidence(x.note, roster, cfg), roster,
                                               " (from the bank-sheet note, not the memo)")
            x.tier, x.fid, x.cands = tier, fid, cands
            x.why = _join(x.why, why)
        if x.prior_ref and x.tier in ("A", "B", "C"):
            x.why += f" (prior-term reference {x.prior_ref}  -  family identified, invoice not loaded)"

    for x in rows:
        single = next(iter(x.cands)) if len(x.cands) == 1 else ""
        x.amt = amount_check(x.amount, x.fid or single, x.term, roster, cfg)
    return rows


def _join(existing: str, extra: str) -> str:
    return f"{existing}; {extra}" if existing else extra
