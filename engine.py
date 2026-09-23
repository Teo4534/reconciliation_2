"""
engine.py - the allocation logic, with no Excel writing and no side effects.

Everything here is a plain function of its inputs, so it can be unit-tested in milliseconds
without generating a workbook. reconcile.py loads the ledger, calls into this module, and
writes the result back out.

Tiers, highest evidence first, first hit wins:
    M  manual override typed into the Receipts sheet                      -> allocated
    A  current-term invoice reference naming exactly one family,
       corroborated: something else in the memo agrees with it (a
       surname, even truncated or misspelt, or a child's name) or the
       payer is already known for that family; or a payer name already
       seen with a verified reference (an alias)                          -> allocated
    B  exactly one roster surname found in the memo                       -> allocated
    C  fuzzy evidence: truncated or misspelt surname, or a child's name;
       or a reference that nothing else in the memo supports              -> review
    D  nothing usable, or several families fit                            -> review
    X  the reference and the memo point to different families, either
       because the memo names other families outright or because another
       family's surname evidence strictly contains the referenced
       family's, so the memo describes that family and more              -> review
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
    # Inclusive (low, high) bounds on the three-digit part of an invoice number. Needed only when a
    # term and the one before it fall in the same year - autumn 2026 issues 2026-5xx where the
    # January 2026 term issued 2026-0xx, and the year cannot separate them. Empty means the year
    # decides alone, which is what a term whose predecessor is in another year relies on.
    cur_series: tuple = ()
    prior_series: tuple = ()

    @property
    def year_cur(self) -> str:
        return self.term_cur[-4:]

    @property
    def year_prior(self) -> str:
        return self.term_prior[-4:]

    @staticmethod
    def series_bounds(text) -> tuple:
        """Inclusive (low, high) from a series written for people, "2026-5xx / 2026-6xx" -> (500, 699).
        The Terms sheet and terms.py both carry the human form; this is the one place it is read."""
        nums = re.findall(r"(\d)xx", str(text or ""))
        return (int(min(nums)) * 100, int(max(nums)) * 100 + 99) if nums else ()

    def in_series(self, num: str, which: str) -> bool:
        """Is this three-digit part inside that term's invoice series? True when none is set."""
        lo_hi = self.cur_series if which == "cur" else self.prior_series
        return not lo_hi or lo_hi[0] <= int(num) <= lo_hi[1]

    def ref_patterns(self):
        """Compiled regexes for a current-term reference (two spellings) and a prior-term one."""
        y, yy, py = self.year_cur, self.year_cur[2:], self.year_prior
        cur = (re.compile(rf"{y}\s*-?\s*(\d{{3}})"), re.compile(rf"(?<!\d){yy}(\d{{3}})(?!\d)"))
        prior = (re.compile(rf"{py}\s*-?\s*(\d{{3}})"),
                 re.compile(r"(?<!\d)2[1-4]\s*-?\s*1\d\d(?!\d)|(?<!\d)2[1-4]\d{3}(?!\d)"))
        return cur, prior


# ----------------------------------------------------------------------------- text helpers

def norm(x) -> str:
    """Upper-case ASCII with accents stripped, so ELODIE and Élodie compare equal."""
    return unicodedata.normalize("NFKD", str(x or "")).encode("ascii", "ignore").decode().upper()


def tokens(text) -> list[str]:
    """Words of three or more letters, stop words dropped. An apostrophe after a single leading
    letter is removed, so M'BOLO and O'BRIEN read MBOLO and OBRIEN as surname_parts() indexes them;
    every other apostrophe stays a word boundary, so ADAM'S yields ADAM and never the surname ADAMS.
    Typographic apostrophes count as apostrophes (norm() would otherwise drop them and glue)."""
    text = norm(re.sub("[\u2018\u2019\u02bc]", "'", str(text or "")))
    text = re.sub(r"\b([A-Z])'(?=[A-Z])", r"\1", text)
    return [t for t in re.findall(r"[A-Z]{3,}", text) if t not in STOP_WORDS]


def surname_parts(sur) -> list[str]:
    """Index keys for a surname: split on hyphens and spaces, keep letters only, drop parts under
    three letters. M'BOLO and O'BRIEN index as MBOLO and OBRIEN, which is what tokens() yields."""
    parts = (re.sub(r"[^A-Z]", "", p) for p in re.split(r"[-\s]", norm(sur)))
    return [p for p in parts if len(p) >= 3]


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
        for part in surname_parts(sur):
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

def hits(text, roster: Roster, cfg: Config) -> list[tuple[str, str, str, str]]:
    """Every (token, part, fid, signal) the text produces against the roster: an exact surname
    part, a unique child's first name, or a fuzzy hit (the token starts a surname part, misspells
    it, or contains it). part is the roster surname part the token matched - the token itself when
    the hit is exact, and '' for a child's first name, which is not surname evidence. The only
    place the text is matched against the roster; evidence() folds this per family."""
    out = []
    sur_parts = roster.sur_parts
    for t in tokens(text):
        if t in roster.sur_index:
            for f in roster.sur_index[t]:
                out.append((t, t, f, f"surname {t}"))
            continue
        if t in roster.unique_first:
            out.append((t, "", roster.unique_first[t], f"child first name {t.title()}"))
        for p in sur_parts:
            if p.startswith(t) and len(t) >= 4:
                for f in roster.sur_index[p]:
                    out.append((t, p, f, f"'{t}' starts {p}"))
            elif len(t) >= 5 and SequenceMatcher(None, t, p).ratio() >= cfg.fuzzy_ratio:
                for f in roster.sur_index[p]:
                    out.append((t, p, f, f"'{t}' ~ {p}"))
            elif len(t) >= 8 and p in t:
                for f in roster.sur_index[p]:
                    out.append((t, p, f, f"{p} inside '{t}'"))
    return out


def by_family(found: list[tuple[str, str, str, str]]) -> dict[str, list[str]]:
    """Fold hits into fid -> signals, in the order the text produced them."""
    ev = defaultdict(list)
    for _t, _part, f, sig in found:
        ev[f].append(sig)
    return ev


def evidence(text, roster: Roster, cfg: Config) -> dict[str, list[str]]:
    """fid -> the signals found in text: exact surname parts, unique child first names, fuzzy hits."""
    return by_family(hits(text, roster, cfg))


def surname_pairs(found: list[tuple[str, str, str, str]]) -> dict[str, set[tuple[str, str]]]:
    """fid -> the (token, roster surname part) pairs behind its surname evidence in this text. A
    child's first name (part '') is left out: it says nothing about a surname. How much of a part
    a token covers is not weighed, so how far the bank truncated a compound surname - PELLET,
    PELLE, PELL - does not decide whether the compound family is seen at all."""
    pairs = defaultdict(set)
    for t, part, f, _sig in found:
        if part:
            pairs[f].add((t, part))
    return pairs


def richer_families(found: list[tuple[str, str, str, str]], cand: str) -> set[str]:
    """Families whose surname evidence strictly contains cand's, so the memo says everything it
    says for cand and more: NAKASHIMA-PELLETIER, hit by NAKASHIMA and by PELLET, over NAKASHIMA,
    hit by NAKASHIMA alone. Equal evidence is not richer (SANDOVAL against SANDOVAL ZABALA, which
    the reference is left to settle), and different words are not richer (SMITH against DUPONT).
    A family the memo offers nothing for has nothing to contain, so a reference with no name
    support is never contradicted this way: that is tier C's business, not this rule's."""
    pairs = surname_pairs(found)
    own = pairs.get(cand, set())
    return {f for f, ps in pairs.items() if f != cand and ps > own} if own else set()


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
    ev: dict = field(default_factory=dict)      # evidence(memo): fid -> signals, computed once in allocate()
    tier: str = ""
    fid: str = ""
    why: str = ""
    cands: set = field(default_factory=set)
    amt: str = ""


def parse_refs(memo: str, cfg: Config) -> tuple[str, str]:
    """Pull a current-term reference (normalised to YYYY-nnn) and, failing that, a prior-term one.

    A number counts for a term only if it is in that term's series as well as its year, so that two
    terms sharing a year stay apart. Where no series is configured the year decides on its own.
    """
    (cur_a, cur_b), (prior_a, prior_b) = cfg.ref_patterns()
    first = lambda pat, which: next((m for m in pat.finditer(memo) if cfg.in_series(m.group(1), which)), None)
    m = first(cur_a, "cur") or first(cur_b, "cur")
    cur_ref = f"{cfg.year_cur}-{m.group(1)}" if m else ""
    pm = first(prior_a, "prior") or prior_b.search(memo)
    prior_ref = pm.group(0) if (pm and not cur_ref) else ""
    return cur_ref, prior_ref


def read_receipts(ws, cfg: Config) -> list[Receipt]:
    rows = []
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 1).value is None:
            break
        g = lambda c: ws.cell(r, c).value
        # Receipts layout, set in build_ledger.py: A date, B amount, C payer, D reference, E term,
        # F family override, G family, H allocated, I tier, J family (engine), K candidates,
        # L why, M amount check, N flags, O full memo.
        memo = g(15) or ""
        cur_ref, prior_ref = parse_refs(norm(memo), cfg)
        rows.append(Receipt(
            r=r, date=g(1), amount=float(g(2) or 0), payer=g(3) or "", ref=g(4) or "", memo=memo,
            cur_ref=cur_ref, prior_ref=prior_ref, term=g(5), override=str(g(6) or "").strip(),
            note=g(14) or "",
        ))
    return rows


# ----------------------------------------------------------------------------- allocation

def allocate(rows: list[Receipt], roster: Roster, cfg: Config) -> list[Receipt]:
    """Assign a tier, family and reason to every receipt. Two passes:

    1. references and overrides, which also teach us payer aliases. A reference allocates only
       when corroborated: the memo carries some evidence for the same family, or the payer is
       already a usable alias for it (seeded, or learned from an earlier row in this pass).
    2. names, for everything still open, using the aliases learned in pass 1
    """
    alias = defaultdict(set)
    for k, v in roster.alias_seed.items():
        alias[k] |= v
    unconfirmed = defaultdict(set)   # payer key -> families whose reference it quoted uncorroborated

    for x in rows:
        found = hits(x.memo, roster, cfg)
        x.ev = by_family(found)
        x.named = families_named(x.ev)
        x.tier = x.fid = x.why = ""
        x.cands = set()
        pk = payer_key(x.payer)
        if x.override:
            x.tier, x.fid, x.why = "M", x.override, "manual override"
            alias[pk].add(x.override)
            continue
        if not x.cur_ref:
            continue
        fams = roster.inv_index.get(x.cur_ref, set())
        if len(fams) > 1 and len(fams & x.named) == 1:
            fams = fams & x.named
        if len(fams) == 1:
            cand = next(iter(fams))
            richer = richer_families(found, cand)
            if x.named and cand not in x.named:
                x.tier, x.cands = "X", x.named | {cand}
                x.why = wrong_number(x.cur_ref, cand, x.named)
            elif richer:
                x.tier, x.cands = "X", richer | {cand}
                x.why = wrong_number(x.cur_ref, cand, richer)
                # The payer still learns cand, exactly as it would have without this rule: what is
                # held back is this row, not the rest of the pass. Every other row then decides on
                # the same aliases as before, so this rule only ever subtracts an allocation.
                alias[pk].add(cand)
            elif cand in x.ev or known_families(alias, pk, cfg) == {cand}:
                x.tier, x.fid, x.why = "A", cand, f"invoice reference {x.cur_ref}"
                alias[pk].add(cand)
            else:
                x.tier, x.cands = "C", {cand}
                x.why = (f"invoice reference {x.cur_ref} points to {cand} but nothing else in the memo "
                         "supports it  -  confirm, and later payments from this payer will follow")
                if x.ev:   # fuzzy evidence for some other family: not a contradiction, but worth seeing
                    x.why += "  -  memo also fits " + "; ".join(
                        f"{f} ({', '.join(sorted(set(x.ev[f])))})" for f in sorted(x.ev))
                unconfirmed[pk].add(cand)
        elif len(fams) > 1:
            x.tier, x.cands, x.why = "D", fams, f"reference {x.cur_ref} is shared by " + ", ".join(sorted(fams))
        else:
            x.why = f"reference {x.cur_ref} not in roster"

    # only keep aliases that are long enough to be distinctive, point at exactly one family, and
    # never quoted another family's reference (a grandparent paying for two families)
    alias = {k: v for k, v in alias.items()
             if known_families(alias, k, cfg) == v and len(v) == 1 and not (unconfirmed.get(k, set()) - v)}

    for x in rows:
        if x.tier:
            continue
        pk = payer_key(x.payer)
        if pk in alias:
            x.tier, x.fid = "A", next(iter(alias[pk]))
            x.why = _join(x.why, "payer name previously seen with a verified reference for this family")
        else:
            tier, fid, cands, why = decide(x.ev, roster)
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


def wrong_number(cur_ref: str, cand: str, against: set) -> str:
    """Tier X's reason: the reference says one family, the memo says these others. One template,
    whether the memo names them outright or describes one of them more fully than the referenced
    family, so both kinds of contradiction read the same in the review queue."""
    return (f"reference {cur_ref} points to {cand} but the memo names " + ", ".join(sorted(against))
            + "  -  parent may have typed the wrong invoice number")


def known_families(alias: dict, pk: str, cfg: Config) -> set:
    """Families a payer key already stands for, or nothing if the key is too short to trust."""
    return set(alias.get(pk, set())) if len(pk) >= cfg.min_alias_len else set()


def _join(existing: str, extra: str) -> str:
    return f"{existing}; {extra}" if existing else extra
