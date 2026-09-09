"""
terms.py - the school's terms, as data. The one place a term is defined.

A term is a period the office invoices for. Everything that varies between them lives here: how
many sessions of each kind the term bills, the invoice series the office issues for it, whether
that register has been supplied, and the dates between which a receipt is assumed to belong to it.

Adding next term is a new entry in TERMS, not an edit to the code that reads it. Running a term is
`build_ledger.py ... --term <ID>`. Sessions and windows are written to the Terms sheet of the
workbook, where the office can correct them without touching Python.
"""
import datetime as dt

TERMS = [
    dict(id="AUT-2025", desc="Autumn 2025 (Sep-Dec)", sat=11, wed=11, series="2025-5xx / 2025-6xx",
         loaded="No", start=dt.date(2025, 8, 1), end=dt.date(2025, 12, 14),
         note="Sessions inferred from receipts (£214.50 = 11 × £19.50). Invoice register not supplied  -  add it to allocate autumn receipts."),
    dict(id="JAN-2026", desc="Term starting January 2026", sat=21, wed=20, series="2026-xxx",
         loaded="Yes", start=dt.date(2025, 12, 15), end=dt.date(2026, 8, 31),
         note="21 = £409.50 ÷ £19.50. Wednesday 20 = £320 ÷ £16 (single club invoice). Confirm both with the office."),
    dict(id="AUT-2026", desc="Autumn 2026 (Sep-Dec)", sat=11, wed=11, series="2026-5xx / 2026-6xx",
         loaded="No", start=dt.date(2026, 9, 1), end=dt.date(2026, 12, 14),
         note="NOT YET CONFIRMED. Sessions copied from the previous autumn and the invoice series assumed to follow it. "
              "Correct both on this sheet, or in terms.py, before running this term for real."),
]

TERM_IDS = [t["id"] for t in TERMS]


def default_term():
    """The most recent term whose invoice register has been supplied.

    A term with no register can have its receipts tied to a family but never to an invoice, so it
    is not a sensible thing to reconcile by default. Mark a register loaded and the default moves.
    """
    loaded = [t for t in TERMS if t["loaded"] == "Yes"]
    return (loaded[-1] if loaded else TERMS[-1])["id"]


def resolve(term_id):
    """Return (prior, current) for a term ID, or exit with a message naming the terms that exist."""
    if term_id not in TERM_IDS:
        raise ValueError(f"unknown term {term_id!r}. terms.py defines: {', '.join(TERM_IDS)}")
    i = TERM_IDS.index(term_id)
    if i == 0:
        raise ValueError(f"{term_id} is the earliest term defined, so it has no prior term to compare against.")
    return TERMS[i - 1], TERMS[i]
