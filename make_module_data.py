"""Generate the extra datasets for Weeks 2-3, 4 and 6 (fixed seed, so every run is identical).

Run once: python make_module_data.py
Writes to data/:
  ED_Barrie_OneDay.csv      30 visits from one day at Barrie (Weeks 2-3, frequency tables)
  Supply_Order.csv          10-line supply order (Weeks 2-3, arithmetic and references)
  ED_Intake_Raw.xlsx        messy intake export + Sites lookup sheet (Week 4, what students get)
  ED_Intake_Clean.csv       the same rows cleaned correctly (Week 4, what Dana checks against)
  ED_Visits_Dates.csv       visit dates, birth dates and follow-up dates (Week 6)
ED_Visits_2025.csv (Week 5) comes from make_dataset.py and is not changed here.
"""
from datetime import date, timedelta

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.styles import Font

rng = np.random.default_rng(2026)


def pick(options, weights):
    w = np.array(weights, dtype=float)
    return options[rng.choice(len(options), p=w / w.sum())]


# ------------------------------------------------------------ Weeks 2-3: one day at Barrie
rows = []
for i in range(30):
    ctas = pick([1, 2, 3, 4, 5], [1, 4, 11, 9, 5])
    base = [5, 16, 48, 75, 95][ctas - 1]
    wait = int(max(1, round(base * np.exp(0.4 * rng.standard_normal()))))
    band = pick([0, 1, 2, 3], [5, 9, 9, 7])
    age = [0, 18, 40, 65][band] + int(rng.integers(0, [18, 22, 25, 28][band]))
    rows.append({"Age": age, "Sex": "F" if rng.random() < 0.5 else "M", "TriageLevel": ctas,
                 "WaitTimeMin": wait})
day = pd.DataFrame(rows)
day.insert(0, "VisitID", [f"BAR-0315-{i + 1:02d}" for i in range(30)])
day.to_csv("data/ED_Barrie_OneDay.csv", index=False)

supplies = pd.DataFrame([
    ("Nitrile gloves (box of 100)", 40, 12.50, 11.75),
    ("Surgical masks (box of 50)", 30, 9.80, 10.40),
    ("IV start kits", 120, 4.15, 3.90),
    ("Saline flush syringes", 400, 0.62, 0.58),
    ("Gauze pads (pack of 100)", 25, 7.40, 7.40),
    ("Alcohol swabs (box of 200)", 35, 3.20, 2.95),
    ("Blood pressure cuffs", 6, 48.00, 52.00),
    ("Thermometer probe covers (box)", 18, 14.90, 13.60),
    ("Suture kits", 60, 11.25, 10.80),
    ("Hand sanitizer (1 L)", 50, 6.75, 7.10),
], columns=["Item", "Quantity", "UnitCostCAD", "LastYearUnitCostCAD"])
supplies.to_csv("data/Supply_Order.csv", index=False)

# ------------------------------------------------------------ Week 4: messy intake export
FIRST = ["Anna", "Ben", "Chloe", "David", "Emma", "Farah", "George", "Hana", "Ian", "Julia", "Kevin",
         "Leah", "Marco", "Nadia", "Omar", "Priya", "Quinn", "Rosa", "Sam", "Tara", "Victor", "Wendy"]
LAST = ["Smith", "Lee", "Patel", "Brown", "Martin", "Nguyen", "Wilson", "Clark", "Singh", "Roy",
        "Walker", "Young", "King", "Scott", "Green", "Baker", "Hill", "Wood"]
SITES = [("BAR", "Barrie", 40), ("ORI", "Orillia", 25), ("MID", "Midland", 18), ("COL", "Collingwood", 17)]

raw, clean = [], []
for i in range(60):
    code, site, _ = SITES[rng.choice(4, p=[0.4, 0.25, 0.18, 0.17])]
    first, last = FIRST[rng.integers(len(FIRST))], LAST[rng.integers(len(LAST))]
    # name: mixed case, sometimes extra spaces
    style = rng.integers(4)
    name = [f"{last.upper()}, {first.upper()}", f"{last.lower()}, {first.lower()}",
            f"{last}, {first}", f"{last.upper()}, {first.lower()}"][style]
    if rng.random() < 0.25:
        name = "  " + name + " "
    # visit reference: site code in the first 3 letters, case varies
    ref = f"{code}-2025-{i + 1:04d}"
    if rng.random() < 0.3:
        ref = ref.lower()
    age = int(rng.integers(19, 92))
    age_raw = str(age) if rng.random() > 0.15 else f" {age}"   # stored as text, some with a space
    if rng.random() < 0.07:
        age_raw, age = "NA", None
    sex = "F" if rng.random() < 0.52 else "M"
    sex_raw = pick([sex, sex.lower(), {"F": "Female", "M": "Male"}[sex], {"F": "female", "M": "male"}[sex]],
                   [55, 15, 20, 10])
    ctas = int(pick([1, 2, 3, 4, 5], [2, 12, 38, 33, 15]))
    older = age is not None and age >= 60
    sys_ = int(round(rng.normal(132 if older else 121, 14)))
    dia = int(round(rng.normal(82 if older else 77, 9)))
    r = rng.random()
    if r < 0.08:
        bp_raw, sys_, dia = "NA", None, None
    elif r < 0.12:
        bp_raw, sys_, dia = "n/a", None, None
    elif r < 0.15:
        bp_raw, sys_, dia = None, None, None
    else:
        bp_raw = f"{sys_}/{dia}"
    wait = int(max(2, round([5, 16, 50, 80, 100][ctas - 1] * np.exp(0.35 * rng.standard_normal()))))
    r = rng.random()
    if r < 0.08:
        wait_raw, wait = "NA", None
    elif r < 0.11:
        wait_raw, wait = None, None
    else:
        wait_raw = wait
    if sys_ is None:
        cat = None
    elif sys_ >= 140 or dia >= 90:
        cat = "High"
    elif sys_ < 120 and dia < 80:
        cat = "Normal"
    else:
        cat = "Elevated"
    vid = f"V{i + 1:03d}"
    raw.append([vid, ref, name, age_raw, sex_raw, ctas, bp_raw, wait_raw])
    clean.append({"VisitID": vid, "SiteCode": code, "SiteName": site, "PatientName": f"{last}, {first}",
                  "Age": age, "Sex": sex, "TriageLevel": ctas, "Systolic": sys_, "Diastolic": dia,
                  "BPCategory": cat, "WaitTimeMin": wait})

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Intake"
ws.append(["VisitID", "VisitRef", "PatientName", "Age", "Sex", "TriageLevel", "BloodPressure", "WaitTimeMin"])
for r in raw:
    ws.append(r)
for c in ws["D"][1:]:          # ages stay text, the way the export system wrote them
    c.data_type = "s"
    c.number_format = "@"
s2 = wb.create_sheet("Sites")
s2.append(["SiteCode", "SiteName", "Beds"])
for code, site, beds in SITES:
    s2.append([code, site, beds])
for sh in (ws, s2):
    for c in sh[1]:
        c.font = Font(bold=True)
    sh.freeze_panes = "A2"
    for col in sh.columns:
        sh.column_dimensions[col[0].column_letter].width = max(len(str(c.value or "")) for c in col) + 3
wb.save("data/ED_Intake_Raw.xlsx")
pd.DataFrame(clean).to_csv("data/ED_Intake_Clean.csv", index=False)

# ------------------------------------------------------------ Week 6: dates
ed = pd.read_csv("data/ED_Visits_2025.csv")
dob, fu = [], []
for v, a in zip(ed.VisitDate, ed.Age):
    vd = date.fromisoformat(v)
    # birthday between a and a+1 years before the visit, so DATEDIF(DOB, visit, "Y") = Age
    latest = date(vd.year - a, vd.month, min(vd.day, 28))
    dob.append(latest - timedelta(days=int(rng.integers(1, 360))))
    fu.append(vd + timedelta(days=int(rng.integers(2, 31))) if rng.random() < 0.7 else None)
dates = pd.DataFrame({"VisitID": ed.VisitID, "VisitDate": ed.VisitDate, "DateOfBirth": [d.isoformat() for d in dob],
                      "Site": ed.Site, "TriageLevel": ed.TriageLevel,
                      "FollowUpDate": [d.isoformat() if d else None for d in fu]})
dates.to_csv("data/ED_Visits_Dates.csv", index=False)
print("written:", len(day), len(supplies), len(raw), len(dates))
