"""
sources.py - what the roster and the bank export look like, as data. The one place a heading is named.

build_ledger.py reads the files through these tables and preflight.py checks them against the same
tables, so the two cannot disagree about what a valid file is. A roster with a new heading, or a
bank that labels a column differently, is one string added here.
"""

# Each canonical roster field lists every heading it has been seen under, most recent first.
ROSTER_COLS = {
    "status":   ("Statut:",),
    "pupil":    ("LES ELEVES",),
    "invoice":  ("INVOICE NUMBER", "INVOICE N"),
    "fees":     ("FEES", "REG FEE"),
    "supplies": ("OFFICE SUPPLIES", "SUPPLIES"),
    "reg":      ("REG FEES ONE OFF", "ONEOFF REG"),
    "paid":     ("PAID",),
}
# Fields the ledger cannot be built without. The rest are optional; PAID is free text.
ROSTER_REQUIRED = ("status", "pupil", "invoice", "fees")

# The bank export is read by heading when it has one and by position when it does not. Barclays
# puts the columns in a different order from the fixture, so position alone is not enough.
BANK_COLS = {
    "date":   ("Date",),
    "amount": ("Amount",),
    "memo":   ("Memo",),
    "cat":    ("Subcategory", "Category"),
}
BANK_POSITIONAL = ["date", "amount", "memo", "cat", "note", "extra"]

# Outgoing lines in these categories are the school's own spending, not a fee receipt. Only
# negative amounts are ever dropped, so no incoming money can be lost to this filter.
NOT_A_RECEIPT = frozenset({"direct debit", "contactless card purchase", "bill payment",
                           "card purchase", "debit", "standing order", "credit payment"})

# Invoice numbers as typed. A question mark is the office saying it has not issued one yet.
INVOICE_BLANKS = frozenset({"?", "??", "-", "n/a", "tbc", "nan"})

# Words the office types beside a name that are not part of it: FRATERIE marks a sibling group,
# ESSAIS a trial. Left in, they become surname evidence the bank memo will never match.
NAME_NOISE = frozenset({"FRATERIE", "FRERE", "SOEUR", "ESSAIS", "ESSAI", "NOUVEAU", "NOUVELLE"})


def bank_header(first_row) -> bool:
    """Does the first row of a bank export carry headings, so columns are read by name?"""
    head = [str(x).strip() for x in first_row]
    return any(h in head for h in BANK_COLS["memo"] + BANK_COLS["amount"])
