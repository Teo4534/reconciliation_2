"""
sources.py - what the roster and the bank export look like, as data. The one place a heading is named.

build_ledger.py reads the files through these tables and preflight.py checks them against the same
tables, so the two cannot disagree about what a valid file is. A roster with a new heading, or a
bank that labels a column differently, is one string added here. Where the roster sits inside a
workbook (which sheet, how many rows above the headings) is found here too, by locate_roster.
"""
import pandas as pd

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


# Rows of furniture a sheet may carry above its headings: teacher, assistant, class name, room.
# The office's master workbook puts the headings on row 5. Anything deeper is not a roster sheet.
HEADER_SCAN_ROWS = 12


def header_row(rows):
    """Index of the first row that names every required roster field, or None.

    `rows` is a sequence of rows, each a sequence of cell values. A cell matches a field when it
    equals one of that field's aliases in ROSTER_COLS, after stripping. Exact, not fuzzy: a sheet
    that says LES ELEVES but has no invoice column is a class list, not the roster.
    """
    for i, row in enumerate(rows):
        cells = {str(v).strip() for v in row if v is not None and str(v).strip() != ""}
        if all(any(alias in cells for alias in ROSTER_COLS[field_]) for field_ in ROSTER_REQUIRED):
            return i
    return None


def locate_roster(path):
    """(sheet name, header row index) of the roster inside a workbook, or None if no sheet has one.

    The office keeps the roster inside its master workbook: a sheet per year, a waiting list, staff,
    and on each year's sheet four rows of teacher and class names above the headings. The first
    real run needed that sheet copied out by hand with the top rows deleted. This finds it instead:
    the first sheet whose top rows contain a heading row naming every required field. Older years'
    sheets carry names but no invoice or fee columns, so they are never chosen. A file that is only
    the roster, headings on row 1, is the trivial case.
    """
    book = pd.ExcelFile(path)
    for sheet in book.sheet_names:
        top = pd.read_excel(book, sheet_name=sheet, header=None, nrows=HEADER_SCAN_ROWS)
        i = header_row(top.itertuples(index=False, name=None))
        if i is not None:
            return sheet, i
    return None


def bank_header(first_row) -> bool:
    """Does the first row of a bank export carry headings, so columns are read by name?"""
    head = [str(x).strip() for x in first_row]
    return any(h in head for h in BANK_COLS["memo"] + BANK_COLS["amount"])
