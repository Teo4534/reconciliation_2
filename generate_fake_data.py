"""
generate_fake_data.py - build a fictional roster and bank export that reproduce the reference
failure modes seen in a real school fee ledger, with a ground-truth answer for every bank line.

    python generate_fake_data.py [outdir] [--seed N]

Writes roster.xlsx, bank.xlsx and ground_truth.csv. No real pupil, parent or payment data is
used anywhere: names are drawn from the lists below and every amount is generated from the
published fee rules. The point is to reproduce the shape of the mess, not any record.

Failure modes reproduced (see FAILURE_MODES for the mix):
  clean            reference exactly as invoiced
  glued            reference run together with the payer's name, no separator
  no_hyphen        year and number with the hyphen dropped, or spaces inserted
  prior_term       last term's reference reused for this term's payment
  bare             just the three-digit invoice number
  name_only        no reference at all, surname in the memo
  first_name_only  no reference, only a child's first name
  wrong_ref        a valid reference belonging to a different family (typo)
  truncated        reference cut off by the bank's fixed-width memo field
  shared_invoice   two unrelated families issued the same invoice number
"""
import csv, random, sys, unicodedata
from pathlib import Path
import pandas as pd

OUTDIR = Path(sys.argv[1]) if len(sys.argv) > 1 and not sys.argv[1].startswith("-") else Path("examples")
SEED = int(sys.argv[sys.argv.index("--seed") + 1]) if "--seed" in sys.argv else 20260101
rng = random.Random(SEED)

RATE = {1: 19.50, 2: 34.50, 3: 46.00, 4: 57.10}
SESSIONS_CUR, SESSIONS_PRIOR = 21, 11
REG_FEE, SUPPLIES = 25.00, 25.00
TERM_CUR, TERM_PRIOR = "2026", "2025"

SURNAMES = """Abarca Ashworth Baptiste Beauchamp Belhadj Bergqvist Blomfield Cardoso Chidozie Comeau
Darrieussecq Delacroix Duplantier Eriksen Falzon Fitzhugh Gaillard Grynberg Halvorsen Hollingworth
Ibarra Jankowski Kaczmarek Kilbride Lachapelle Lindqvist Maalouf Marchetti Nakashima Nwachukwu
Okonkwo Pelletier Perreault Quilliam Rasmussen Rocheteau Sandoval Sigurdsson Thibodeaux Trevelyan
Uchendu Vanterpool Villalobos Wendelin Wojciechowski Yankovic Zabala Zimmerli Achterberg Bosanquet
Auclair Brannigan Castellane Dziedzic Ferreiro Grimaldsen Hasanovic Ivarsson Joubert Kristoffersen
Lindegaard Montenegro Nordstrom Oyelaran Pietrangeli Quintanilla Rothermere Saltzman Tsvetkova Ulvaeus""".split()
DOUBLE = [("Achterberg", "Marchetti"), ("Belhadj", "Trevelyan"), ("Nakashima", "Pelletier")]
FIRST = """Amara Anouk Bastien Beatriz Cato Celestine Dario Eleni Emeka Esben Fabienne Gaspard Halima
Ilaria Ingvar Jomo Juniper Kaia Kwame Leocadie Lorcan Maelys Matteo Nadia Ndidi Olamide Otto Piero
Quentin Rafaela Rune Sable Solveig Tamsin Thibault Uma Valentin Wren Xanthe Yusuf Zephyr Zosia
Anselm Bijou Caspian Delphine Ezra Fionn Giulia Hektor""".split()
PARENT_FIRST = """Adrien Bettina Camille Dermot Elodie Fergus Genevieve Hugo Ingrid Jerome Katrin
Ludovic Margit Nils Orla Pascal Rosalind Stefan Tove Ulrich Vera Wilhelm Yolande Zora""".split()

FAILURE_MODES = [("clean", 14), ("glued", 6), ("no_hyphen", 5), ("prior_term", 9), ("bare", 4),
                 ("name_only", 9), ("first_name_only", 3), ("wrong_ref", 3), ("truncated", 5), ("shared_invoice", 2)]


def strip_accents(x):
    return unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode()


families, used_first = [], set()
n_families = 60
assert n_families - len(DOUBLE) <= len(SURNAMES), "not enough distinct surnames"
for i in range(n_families):
    if i < len(DOUBLE):
        sur = "-".join(DOUBLE[i])
    else:
        sur = SURNAMES[i - len(DOUBLE)]
    n_kids = rng.choices([1, 2, 3, 4], weights=[52, 33, 12, 3])[0]
    kids = []
    for _ in range(n_kids):
        f = rng.choice([x for x in FIRST if x not in used_first] or FIRST)
        used_first.add(f)
        kids.append(f)
    families.append(dict(sur=sur, kids=kids, parent=rng.choice(PARENT_FIRST)))

inv_numbers = {}
n = 1
for i, fam in enumerate(families):
    inv_numbers[i] = f"{TERM_CUR}-{n:03d}"
    n += 1
shared_a, shared_b = 7, 34
inv_numbers[shared_b] = inv_numbers[shared_a]

roster_rows = []
for i, fam in enumerate(families):
    k = len(fam["kids"])
    per_child_tuition = round(RATE[min(k, 4)] * SESSIONS_CUR / k, 2)
    for j, kid in enumerate(fam["kids"]):
        roster_rows.append({
            # Which family the row belongs to. Not a roster column: the DataFrame below is built
            # with an explicit column list, so this is dropped before anything is written out. It
            # exists because two families can share an invoice number, so the number cannot be
            # used to group a family's rows.
            "_fam": i,
            "Statut:": "Inscrit",
            "LES ELEVES": f"{fam['sur'].upper()} {kid}",
            "INVOICE NUMBER": inv_numbers[i] + ("a" if j == 0 and k > 1 else ""),
            "FEES": per_child_tuition,
            "OFFICE SUPPLIES": SUPPLIES if rng.random() < 0.85 else None,
            "REG FEES ONE OFF": REG_FEE if rng.random() < 0.18 else None,
            "PAID": rng.choice(["paid", "paid", "part paid", None]),
        })
roster_rows[3]["OFFICE SUPPLIES"] = 18.0
roster_rows[11]["OFFICE SUPPLIES"] = -30.0
roster_rows[19]["INVOICE NUMBER"] = ""
roster_rows[19]["FEES"] = None
roster = pd.DataFrame(roster_rows,
                      columns=["Statut:", "LES ELEVES", "INVOICE NUMBER", "FEES", "OFFICE SUPPLIES", "REG FEES ONE OFF", "PAID"])


def memo(payer, ref):
    return f"{payer[:23]:<23}\t{ref[:21]}\tBGC\t"


def expected_total(i):
    """What family i was invoiced: its own children's rows only.

    Grouping by invoice number instead would bill one family for two, because shared_a and
    shared_b are deliberately issued the same number.
    """
    rows = [r for r in roster_rows if r["_fam"] == i]
    tuition = sum(r["FEES"] or 0 for r in rows)
    extras = sum((r["OFFICE SUPPLIES"] or 0) + (r["REG FEES ONE OFF"] or 0) for r in rows)
    return round(tuition + extras, 2)


modes = [m for m, w in FAILURE_MODES for _ in range(w)]
rng.shuffle(modes)
bank, truth = [], []
TERM_START = pd.Timestamp("2026-01-06")
def has_invoice(i):
    """Was family i actually issued an invoice number this term?

    One child is deliberately left with no invoice number and no fee, to reproduce a roster the
    office really did send. A family in that state cannot quote a reference it was never given,
    so it does not pay: otherwise the generator emits a 'clean' reference that is not in the
    roster, which is a different failure mode than the one it is labelled with.
    """
    return any(str(r["INVOICE NUMBER"]).strip() for r in roster_rows if r["_fam"] == i)


payers = [i for i in range(n_families) if expected_total(i) > 0 and has_invoice(i)]
for idx, i in enumerate(rng.sample(payers, min(len(modes), len(payers)))):
    mode = modes[idx]
    part = rng.random() < 0.30
    if mode == "shared_invoice":
        # Two unrelated families were issued the same invoice number. Pick one of them and let it
        # pay its own invoice, so the payer's name is genuinely theirs and only the number is
        # ambiguous. The family has to be chosen before anything is derived from it: choosing
        # afterwards would leave the payer, the amount and the answer key belonging to a third
        # family, which is the wrong_ref case wearing this label.
        i = rng.choice([shared_a, shared_b])
    fam = families[i]
    payer = strip_accents(f"{fam['parent']} {fam['sur']}".upper())
    ref_clean = inv_numbers[i]
    total = expected_total(i)
    amount = round(total / 2, 2) if part else total

    if mode == "clean":
        ref = ref_clean
    elif mode == "glued":
        ref = f"{fam['sur'][:6].upper()}{ref_clean}"
    elif mode == "no_hyphen":
        ref = rng.choice([ref_clean.replace("-", ""), ref_clean.replace("-", " - ")])
    elif mode == "prior_term":
        ref = f"{TERM_PRIOR}-{rng.randint(500, 660):03d}"
    elif mode == "bare":
        ref = ref_clean.split("-")[1]
    elif mode == "name_only":
        ref = rng.choice([fam["sur"], f"{fam['sur']} fees", "school fees"])
    elif mode == "first_name_only":
        ref = rng.choice(fam["kids"])
    elif mode == "wrong_ref":
        ref = inv_numbers[(i + 1) % n_families]
    elif mode == "truncated":
        ref = f"{fam['sur'].upper()} FRENCH SCHOOL {ref_clean}"[:21]
    elif mode == "shared_invoice":
        ref = ref_clean          # correct for this family, and also correct for another one

    day = TERM_START + pd.Timedelta(days=rng.randint(0, 120))
    bank.append([day, amount, memo(payer, ref), "Registration fees", None, None])
    truth.append(dict(date=day.date(), amount=amount, payer=payer, reference=ref, failure_mode=mode,
                      true_family=fam["sur"], true_invoice=ref_clean))
    if part:
        day2 = day + pd.Timedelta(days=rng.randint(10, 45))
        ref2 = rng.choice([fam["sur"], "2nd payment", ref_clean, f"{fam['sur']} 2nd"])
        bank.append([day2, round(total - amount, 2), memo(payer, ref2), "Registration fees", None, None])
        truth.append(dict(date=day2.date(), amount=round(total - amount, 2), payer=payer, reference=ref2,
                          failure_mode=mode + "+instalment", true_family=fam["sur"], true_invoice=ref_clean))

bank.sort(key=lambda r: r[0])
bank.append([pd.Timestamp("2026-03-11"), -RATE[1] * SESSIONS_CUR, memo("REFUND", "withdrawal refund"), "Registration fees", "refund", None])
truth.append(dict(date="2026-03-11", amount=-RATE[1] * SESSIONS_CUR, payer="REFUND", reference="withdrawal refund",
                  failure_mode="outflow", true_family="", true_invoice=""))

OUTDIR.mkdir(parents=True, exist_ok=True)
roster.to_excel(OUTDIR / "roster.xlsx", index=False)
pd.DataFrame(bank, columns=["date", "amount", "memo", "category", "note", "extra"]).to_excel(
    OUTDIR / "bank.xlsx", index=False, header=False)
with open(OUTDIR / "ground_truth.csv", "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(truth[0]))
    w.writeheader()
    w.writerows(truth)

print(f"seed {SEED}: {len(roster)} pupils in {n_families} families, {len(bank)} bank lines -> {OUTDIR}/")
