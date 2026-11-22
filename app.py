"""Analyst Inbox: students chat with Dana Morgan, a pretend Decision Support manager.

Runs on Streamlit. Needs these secrets (Streamlit Cloud > App > Settings > Secrets):
    ANTHROPIC_API_KEY = "sk-ant-..."
    CLASS_CODE = "something-you-give-students"   # keeps strangers from using your credits
    MODEL = "claude-sonnet-5-5"                   # optional; "claude-haiku-5-5" is cheaper
"""
import base64
import difflib
import io
import json
import re
import secrets
import time
from datetime import datetime
from pathlib import Path

import anthropic
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

SQL_APP = False
DATA_DIR = Path(__file__).parent / "data"
MAX_STUDENT_MESSAGES = 60      # per session; caps cost if someone leaves a script running
MAX_WORKBOOK_CHARS = 20_000    # how much of an uploaded workbook Dana reads
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

st.set_page_config(page_title="Analyst Inbox", page_icon="📨", layout="wide")


# ---------------------------------------------------------------- datasets
# Each dataset has a "check" table Dana computes on, and a file the analyst gets. For most datasets the file
# is the check table minus Dana's check-only columns; for the messy Week 4 data the file is the raw export.
def _dates_check(df: pd.DataFrame) -> pd.DataFrame:
    v, b, f = (pd.to_datetime(df[c]) for c in ("VisitDate", "DateOfBirth", "FollowUpDate"))
    age = v.dt.year - b.dt.year - ((v.dt.month < b.dt.month) | ((v.dt.month == b.dt.month) & (v.dt.day < b.dt.day)))
    df = df.copy()
    df["AgeAtVisit"] = age                                   # = DATEDIF(DateOfBirth, VisitDate, "Y")
    df["DaysToFollowUp"] = (f - v).dt.days                   # = FollowUpDate - VisitDate (blank if no follow-up)
    df["VisitWeekday"] = v.dt.dayofweek.map(dict(enumerate(WEEKDAYS)))   # = TEXT(VisitDate, "ddd")
    df["VisitWeekdayNum"] = (v.dt.dayofweek + 1) % 7 + 1     # = WEEKDAY(VisitDate) (1 = Sunday)
    df["VisitMonthNum"] = v.dt.month                         # = MONTH(VisitDate)
    df["FollowUpBooked"] = f.notna().map({True: "Yes", False: "No"})
    return df


def _supply_check(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["LineTotalCAD"] = (df.Quantity * df.UnitCostCAD).round(2)
    df["HSTCAD"] = (df.LineTotalCAD * 0.13).round(2)
    df["LineTotalWithHSTCAD"] = (df.LineTotalCAD * 1.13).round(2)
    df["UnitCostChangePct"] = ((df.UnitCostCAD - df.LastYearUnitCostCAD) / df.LastYearUnitCostCAD * 100).round(2)
    return df


@st.cache_data
def load_datasets() -> dict:
    ed = pd.read_csv(DATA_DIR / "ED_Visits_2025.csv")
    day = pd.read_csv(DATA_DIR / "ED_Barrie_OneDay.csv")
    sup = pd.read_csv(DATA_DIR / "Supply_Order.csv")
    intake = pd.read_csv(DATA_DIR / "ED_Intake_Clean.csv")
    dates = pd.read_csv(DATA_DIR / "ED_Visits_Dates.csv")
    return {
        "ED_Barrie_OneDay": {"check": day, "file_cols": list(day.columns), "date_cols": [],
                             "numeric": ["Age", "TriageLevel", "WaitTimeMin"]},
        "Supply_Order": {"check": _supply_check(sup), "file_cols": list(sup.columns), "date_cols": [],
                         "numeric": ["Quantity", "UnitCostCAD", "LastYearUnitCostCAD", "LineTotalCAD", "HSTCAD",
                                     "LineTotalWithHSTCAD", "UnitCostChangePct"]},
        "ED_Intake_Raw": {"check": intake, "file_path": DATA_DIR / "ED_Intake_Raw.xlsx",
                          "file_cols": ["VisitID", "VisitRef", "PatientName", "Age", "Sex", "TriageLevel",
                                        "BloodPressure", "WaitTimeMin"], "date_cols": [],
                          "numeric": ["Age", "TriageLevel", "Systolic", "Diastolic", "WaitTimeMin"]},
        "ED_Visits_2025": {"check": ed, "file_cols": list(ed.columns), "date_cols": ["VisitDate"],
                           "numeric": ["Age", "TriageLevel", "WaitTimeMin", "LengthOfStayHrs", "CostCAD"]},
        "ED_Visits_Dates": {"check": _dates_check(dates), "file_cols": list(dates.columns),
                            "date_cols": ["VisitDate", "DateOfBirth", "FollowUpDate"],
                            "numeric": ["TriageLevel", "AgeAtVisit", "DaysToFollowUp", "VisitWeekdayNum",
                                        "VisitMonthNum"]},
    }


DATASETS = load_datasets()
ALL_FILE_COLUMNS = {c for d in DATASETS.values() for c in d["file_cols"]}


def dataset(name: str | None) -> dict:
    name = name or MODULES[st.session_state.topic]["datasets"][0]
    if name not in DATASETS:
        raise ValueError(f"Unknown dataset {name}. Use one of: {', '.join(MODULES[st.session_state.topic]['datasets'])}")
    return DATASETS[name]


@st.cache_data
def raw_messy_facts() -> str:
    """Exact counts of the mess planted in the Week 4 raw export, so Dana can check cleaning work."""
    ws = openpyxl.load_workbook(DATA_DIR / "ED_Intake_Raw.xlsx")["Intake"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    col = {h: i for i, h in enumerate(next(ws.iter_rows(max_row=1, values_only=True)))}

    def tally(name):
        out = {}
        for r in rows:
            v = r[col[name]]
            is_reading = isinstance(v, (int, float)) or (isinstance(v, str) and re.fullmatch(r"\s*\d+(/\d+)?\s*", v))
            k = "(blank)" if v is None else (None if is_reading else v)
            if k is not None:
                out[k] = out.get(k, 0) + 1
        return out

    sex = {}
    for r in rows:
        sex[r[col["Sex"]]] = sex.get(r[col["Sex"]], 0) + 1
    spaced = sum(1 for r in rows if r[col["PatientName"]] != r[col["PatientName"]].strip())
    spaced_age = sum(1 for r in rows if isinstance(r[col["Age"]], str) and r[col["Age"]] != r[col["Age"]].strip())
    lower_ref = sum(1 for r in rows if r[col["VisitRef"]][:3].islower())
    wait_na = tally("WaitTimeMin")
    return (f"Raw export facts (sheet Intake, {len(rows)} rows, columns A-H = {', '.join(col)}): "
            f"BloodPressure non-readings {tally('BloodPressure')}; WaitTimeMin non-numbers {wait_na}; "
            f"Age is stored as TEXT in every row (needs VALUE), {spaced_age} ages have a leading space, "
            f"Age non-numbers {tally('Age')}; "
            f"Sex spellings {sex}; {spaced} PatientName values have extra leading/trailing spaces and names come in "
            f"mixed case (PROPER fixes them); {lower_ref} VisitRef values start with a lower-case site code "
            f"(the first 3 letters of VisitRef are the SiteCode: UPPER(LEFT(...,3)), then XLOOKUP on sheet Sites: "
            f"SiteCode, SiteName, Beds). COUNTIF is not case-sensitive, so COUNTIF(G:G,\"NA\") counts NA but not n/a. "
            f"Note: AVERAGE skips text and blanks, so the average wait of the raw column equals the clean average.")


# ---------------------------------------------------------------- data helpers
def apply_filters(df: pd.DataFrame, filters: list | None, numeric: list) -> pd.DataFrame:
    for f in filters or []:
        col, op, val = f["column"], f["op"], f.get("value", "")
        if col not in df.columns:
            raise ValueError(f"Unknown column {col}. Columns: {', '.join(df.columns)}")
        if op == "is_blank":
            df = df[df[col].isna()]
            continue
        if op == "not_blank":
            df = df[df[col].notna()]
            continue
        if col in numeric:
            s, val = df[col], float(val)
        else:
            s, val = df[col].astype(str).str.lower(), str(val).lower()
        mask = {"=": s == val, "!=": s != val, ">": s > val, "<": s < val,
                ">=": s >= val, "<=": s <= val}[op]
        df = df[mask & df[col].notna()]
    return df


def query_data(metric: str, dataset_name: str | None = None, column: str | None = None, filters: list | None = None,
               group_by: str | None = None, group_by2: str | None = None, bins: list | None = None) -> dict:
    """Compute a statistic the same way Excel does (VAR.S, STDEV.S, MEDIAN, QUARTILE.INC). Blanks are skipped."""
    ds = dataset(dataset_name)
    df = apply_filters(ds["check"], filters, ds["numeric"])
    if metric != "count" and column not in ds["numeric"]:
        raise ValueError(f"{metric} needs a numeric column: {', '.join(ds['numeric'])}")

    def calc(d):
        if metric == "count":
            return int(len(d) if column is None else d[column].notna().sum())
        s = d[column].dropna()
        if s.empty:
            return None
        v = {"sum": s.sum(), "mean": s.mean(), "median": s.median(), "min": s.min(), "max": s.max(),
             "range": s.max() - s.min(), "var": s.var(ddof=1), "stdev": s.std(ddof=1),
             "q1": s.quantile(0.25), "q3": s.quantile(0.75),
             "iqr": s.quantile(0.75) - s.quantile(0.25)}[metric]
        return None if pd.isna(v) else round(float(v), 4)

    def keys(col, d):
        if col not in ds["check"].columns:
            raise ValueError(f"Unknown group_by column {col}")
        present = d[col].dropna().unique()
        order = MONTHS if col == "Month" else WEEKDAYS if col == "VisitWeekday" else sorted(present, key=str)
        return [k for k in order if k in set(present)]

    def py(k):
        return k if isinstance(k, str) else int(k) if float(k).is_integer() else float(k)

    if bins:
        if column not in ds["numeric"]:
            raise ValueError("bins needs a numeric column")
        edges = sorted(float(b) for b in bins)
        s = df[column].dropna()
        out = [{"from": py(lo), "to_below": py(hi) if hi is not None else None,
                "n": int(((s >= lo) & ((s < hi) if hi is not None else True)).sum())}
               for lo, hi in zip(edges, edges[1:] + [None])]
        return {"rows_matched": int(len(s)), "below_first_bin": int((s < edges[0]).sum()), "bins": out}
    if not group_by:
        return {"rows_matched": int(len(df)), "result": calc(df)}
    groups = []
    for k in keys(group_by, df):
        g = df[df[group_by] == k]
        item = {group_by: py(k), "n": int(len(g))}
        if group_by2:
            item["by_" + group_by2] = [{group_by2: py(k2), "result": calc(g[g[group_by2] == k2])}
                                       for k2 in keys(group_by2, df)]
        else:
            item["result"] = calc(g)
        groups.append(item)
    return {"rows_matched": int(len(df)), "groups": groups}


def to_xlsx(df: pd.DataFrame, sheet: str = "Data", date_cols: list | None = None) -> bytes:
    """Excel file with a bold, frozen header row, real dates and sensible column widths."""
    df = df.copy()
    for c in date_cols or []:
        df[c] = pd.to_datetime(df[c]).dt.date
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", date_format="yyyy-mm-dd") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet)
        ws = xw.sheets[sheet]
        ws.freeze_panes = "A2"
        for i, col in enumerate(df.columns, 1):
            ws.cell(1, i).font = openpyxl.styles.Font(bold=True)
            if col in (date_cols or []):
                for (cell,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                    cell.number_format = "yyyy-mm-dd"
            width = max(len(str(col)), *(len(str(v)) for v in df[col].head(200))) + 2
            ws.column_dimensions[get_column_letter(i)].width = min(width, 30)
    return buf.getvalue()


def dataset_file(name: str, filters: list | None = None, columns: list | None = None) -> bytes:
    ds = DATASETS[name]
    if "file_path" in ds:
        if filters or columns:
            raise ValueError("The raw intake export can only be attached whole (no filters or columns).")
        return ds["file_path"].read_bytes()
    df = apply_filters(ds["check"], filters, ds["numeric"])[ds["file_cols"]]
    if columns:
        missing = [c for c in columns if c not in ds["file_cols"]]
        if missing:
            raise ValueError(f"Unknown columns: {missing}")
        df = df[columns]
    if df.empty:
        raise ValueError("No rows match those filters, so there's nothing to attach.")
    return to_xlsx(df, date_cols=[c for c in ds["date_cols"] if c in df.columns])


def attach_dataset(filename: str, dataset_name: str | None = None, filters: list | None = None,
                   columns: list | None = None) -> dict:
    """Dana attaches a dataset (or an extract of it) to her message as an .xlsx file."""
    name = dataset_name or MODULES[st.session_state.topic]["datasets"][0]
    if name not in MODULES[st.session_state.topic]["datasets"]:
        raise ValueError(f"{name} isn't used this module. Use one of: "
                         f"{', '.join(MODULES[st.session_state.topic]['datasets'])}")
    data = dataset_file(name, filters, columns)
    fname = re.sub(r"[^A-Za-z0-9_\-]+", "_", filename.removesuffix(".xlsx")).strip("_") or name
    st.session_state.outbox.append({"name": f"{fname}.xlsx", "data": data, "mime": XLSX_MIME})
    return {"attached": f"{fname}.xlsx", "dataset": name}


def fmt_result(r) -> str:
    if r is None:
        return "(blank, or not calculated)"
    return "(blank)" if r == "" else str(r)


def read_workbook(data: bytes, filename: str) -> str:
    """Turn a student's workbook into text Dana can review: values, and formulas with their results."""
    try:
        wb_f = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
        wb_v = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        return f"[The analyst attached {filename}, but it couldn't be opened. Ask them to re-save it as .xlsx.]"
    out = [f"[The analyst attached the workbook {filename}. Sheets: {', '.join(wb_f.sheetnames)}.]"]
    for ws_f in wb_f.worksheets:
        ws_v = wb_v[ws_f.title]
        rows, cols = ws_f.max_row, ws_f.max_column
        out.append(f"\n--- Sheet '{ws_f.title}' ({rows} rows x {cols} columns) ---")
        if getattr(ws_f, "_pivots", None):
            out.append("(This sheet contains a PivotTable; its values are shown as cells below.)")
        if getattr(ws_f, "_charts", None):
            out.append(f"(This sheet has {len(ws_f._charts)} chart(s); you can't see charts in the file, "
                       "so ask for a screenshot if you need to review one.)")
        # A raw copy of the dataset (its header matches the dataset's columns) is too big to show in
        # full, so show its header and any formulas only. Every other sheet, including PivotTable
        # sheets, is shown in full.
        header = [str(c.value) for c in ws_f[1] if c.value is not None]
        has_pivot = bool(getattr(ws_f, "_pivots", None))
        big = (not has_pivot and rows > 60
               and len(set(header) & ALL_FILE_COLUMNS) >= min(5, len(header)) and len(header) >= 3)
        if big:
            out.append(f"Raw data sheet ({rows - 1} data rows). Header row: {', '.join(header)}. "
                       "Raw values aren't shown. Formula columns are summarised: the first few formulas, "
                       "then a summary of every result in that column.")
        by_col = {}
        for row in ws_f.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                val = c.value
                # Dynamic-array formulas (FILTER, UNIQUE, SORT...) load as ArrayFormula objects.
                if hasattr(val, "text"):
                    val = val.text or ""
                    if not val.startswith("="):
                        val = "=" + val
                if isinstance(val, str):
                    # Strip Excel's internal prefixes so Dana sees =STDEV.S(FILTER(...)), not _xlfn._xlws.
                    val = val.replace("_xlfn._xlws.", "").replace("_xlfn.", "").replace("_xlws.", "")
                is_formula = isinstance(val, str) and val.startswith("=")
                result = ws_v[c.coordinate].value if is_formula else None
                if big:
                    if is_formula:
                        by_col.setdefault(c.column_letter, []).append((c.coordinate, val, result))
                    continue
                if is_formula:
                    out.append(f"{c.coordinate}: {val}  -> {fmt_result(result)}")
                else:
                    out.append(f"{c.coordinate}: {val}")
        for letter, cells in by_col.items():
            head = ws_f[f"{letter}1"].value
            out.append(f"Column {letter}" + (f" ('{head}')" if head is not None else "") + f": {len(cells)} formula cells")
            for coord, f_, r in cells[:3]:
                out.append(f"  {coord}: {f_}  -> {fmt_result(r)}")
            if len(cells) > 3:
                res = [r for _, _, r in cells]
                nums = [r for r in res if isinstance(r, (int, float)) and not isinstance(r, bool)]
                texts = {}
                for r in res:
                    if isinstance(r, str) or r is None:
                        key = "(empty)" if r in (None, "") else r
                        texts[key] = texts.get(key, 0) + 1
                summary = f"  All {len(res)} results: {len(nums)} numbers"
                if nums:
                    summary += (f" (sum {round(sum(nums), 4)}, mean {round(sum(nums) / len(nums), 4)}, "
                                f"min {min(nums)}, max {max(nums)})")
                if texts:
                    top = sorted(texts.items(), key=lambda kv: -kv[1])[:8]
                    summary += "; text/empty results: " + ", ".join(f"{k!s} x{v}" for k, v in top)
                patterns = {re.sub(r"(?<![A-Za-z])([A-Z]{1,3})\$?\d+", r"\1#", f_) for _, f_, _ in cells}
                if len(patterns) > 1:
                    summary += f"; {len(patterns)} different formula patterns in this column"
                out.append(summary)
        text = "\n".join(out)
        if len(text) > MAX_WORKBOOK_CHARS:
            return text[:MAX_WORKBOOK_CHARS] + "\n[Workbook truncated: too large to show in full.]"
    return "\n".join(out)


FILTER_SCHEMA = {"type": "array", "items": {"type": "object", "properties": {
    "column": {"type": "string"},
    "op": {"type": "string", "enum": ["=", "!=", ">", "<", ">=", "<=", "is_blank", "not_blank"]},
    "value": {"type": ["string", "number"], "description": "Not needed for is_blank / not_blank"}},
    "required": ["column", "op"]}}

TOOLS = [
    {
        "name": "query_data",
        "description": "Compute a statistic on one of this module's datasets, the same way Excel would "
                       "(var = VAR.S, stdev = STDEV.S, median = MEDIAN, q1/q3 = QUARTILE.INC; blanks are skipped). "
                       "Use it to check every number the analyst reports before commenting. "
                       "group_by gives one result per category; add group_by2 for a crosstab (like COUNTIFS on two "
                       "columns). bins gives a grouped frequency table: pass the lower bound of each class, e.g. "
                       "[0, 30, 60, 90] for 0-29, 30-59, 60-89, 90 and over.",
        "input_schema": {
            "type": "object",
            "properties": {
                "dataset_name": {"type": "string", "enum": list(DATASETS),
                                 "description": "Defaults to this module's first dataset"},
                "metric": {"type": "string", "enum": ["count", "sum", "mean", "median", "min", "max", "range",
                                                      "var", "stdev", "q1", "q3", "iqr"]},
                "column": {"type": "string", "description": "Numeric column (count needs none; with a column, "
                                                            "count = non-blank cells, like COUNT)"},
                "filters": FILTER_SCHEMA,
                "group_by": {"type": "string"},
                "group_by2": {"type": "string"},
                "bins": {"type": "array", "items": {"type": "number"}},
            },
            "required": ["metric"],
        },
    },
    {
        "name": "attach_dataset",
        "description": "Attach an Excel file to your message: one of this module's datasets, or an extract "
                       "filtered to the rows (and optionally columns) a request needs. The analyst downloads it "
                       "from your message. Check-only columns are never included in the file.",
        "input_schema": {
            "type": "object",
            "properties": {
                "filename": {"type": "string", "description": "Realistic work file name, e.g. ED_Barrie_15Mar"},
                "dataset_name": {"type": "string", "enum": list(DATASETS)},
                "filters": FILTER_SCHEMA,
                "columns": {"type": "array", "items": {"type": "string"}, "description": "Optional subset of columns, in order"},
            },
            "required": ["filename", "dataset_name"],
        },
    },
]

OPEN_REQUEST_TOOL = {
    "name": "open_request",
    "description": "Call this every time you send a NEW request (including the first). It shows the analyst a "
                   "prediction box with your prediction question, and the app won't let them send results for "
                   "this request until they've answered it.",
    "input_schema": {"type": "object", "properties": {
        "title": {"type": "string", "description": "Short name for the request, e.g. 'Barrie top 10 waits'"},
        "prediction_question": {"type": "string", "description": "The question they answer before starting"}},
        "required": ["title", "prediction_question"]},
}


def open_request(title: str, prediction_question: str) -> dict:
    st.session_state.request = {"title": title, "question": prediction_question, "prediction": None}
    return {"ok": True, "note": "The analyst now sees your prediction question in a box and must answer it "
                                "before they can send results for this request."}

TOOLS.append(OPEN_REQUEST_TOOL)
TOOL_FUNCS = {"query_data": query_data, "attach_dataset": attach_dataset, "open_request": open_request}

# ---------------------------------------------------------------- persona
DATASET_DESC = {
    "ED_Barrie_OneDay": (
        "ED_Barrie_OneDay: 30 rows, every visit to the Barrie ED on 15 March 2025 (fictional). Columns: VisitID, "
        "Age, Sex (F, M), TriageLevel (CTAS 1-5; 1 = most urgent), WaitTimeMin (minutes from triage to physician)."),
    "Supply_Order": (
        "Supply_Order: 10 rows, the ED's monthly supply order. Columns: Item, Quantity, UnitCostCAD (this year's "
        "price), LastYearUnitCostCAD. Check-only columns (not in the analyst's file, for checking their formulas): "
        "LineTotalCAD = Quantity x UnitCostCAD, HSTCAD = 13% of the line total, LineTotalWithHSTCAD, "
        "UnitCostChangePct = (this year - last year) / last year x 100. Ontario HST is 13%."),
    "ED_Intake_Raw": (
        "ED_Intake_Raw: the analyst's file is a raw export with two sheets. Sheet 'Intake' (60 rows): VisitID, "
        "VisitRef (e.g. BAR-2025-0012; first 3 letters = site code, case varies), PatientName (\"LAST, FIRST\", "
        "mixed case, some extra spaces), Age (stored as text; some NA), Sex (F, f, Female, male...), TriageLevel, "
        "BloodPressure (text like 120/80; some NA, n/a or blank), WaitTimeMin (some NA or blank). Sheet 'Sites': "
        "SiteCode, SiteName, Beds. You check against the CLEANED version, which has columns VisitID, SiteCode, "
        "SiteName, PatientName (proper case, trimmed), Age (number, blank if NA), Sex (F/M), TriageLevel, Systolic, "
        "Diastolic (numbers from BloodPressure, blank if missing), BPCategory (High if Systolic >= 140 OR "
        "Diastolic >= 90; Normal if Systolic < 120 AND Diastolic < 80; otherwise Elevated; blank if no reading), "
        "WaitTimeMin (blank if NA)."),
    "ED_Visits_2025": (
        "ED_Visits_2025: 300 rows, one per emergency department visit in 2025 (fictional): VisitID, VisitDate, "
        "Month (Jan..Dec), Site (Barrie, Orillia, Midland, Collingwood), Age, AgeGroup (0-17, 18-39, 40-64, 65+), "
        "Sex (F, M), TriageLevel (CTAS 1-5; 1 = most urgent), ChiefComplaint (Abdominal pain, Chest pain, "
        "Shortness of breath, Injury/Fracture, Fever, Headache, Laceration, Mental health), WaitTimeMin, "
        "LengthOfStayHrs, Disposition (Discharged, Admitted, LWBS = left without being seen, Transferred), CostCAD."),
    "ED_Visits_Dates": (
        "ED_Visits_Dates: 300 rows (the same 2025 visits): VisitID, VisitDate, DateOfBirth, Site, TriageLevel, "
        "FollowUpDate (blank when no follow-up was booked; about 70% have one). All dates are real Excel dates. "
        "Check-only columns: AgeAtVisit (= DATEDIF(DateOfBirth, VisitDate, \"Y\")), DaysToFollowUp (= FollowUpDate "
        "- VisitDate, blank if none), VisitWeekday (Mon..Sun, = TEXT(VisitDate,\"ddd\")), VisitWeekdayNum "
        "(= WEEKDAY(VisitDate), 1 = Sunday), VisitMonthNum (= MONTH(VisitDate)), FollowUpBooked (Yes/No)."),
}

MODULES = {
    "Weeks 2–3 · Arithmetic and frequency tables": {
        "datasets": ["ED_Barrie_OneDay", "Supply_Order"],
        "focus": (
            "Basic Excel arithmetic: + - * /, percentages, number formats (currency, percent, decimal places), and "
            "relative vs absolute references ($A$1, and mixed references like $A1 or A$1). Then frequency "
            "distributions built WITHOUT counting functions (no COUNTIF, COUNTIFS, FREQUENCY or PivotTables): "
            "the analyst sorts the small dataset and counts by hand, then uses formulas for relative frequency "
            "(f / total), cumulative frequency, and relative cumulative frequency. Also a grouped frequency table, "
            "e.g. WaitTimeMin in classes of 30 minutes. Use Supply_Order for the arithmetic requests (line totals, "
            "HST at 13% kept in ONE cell and referenced absolutely, percent of total order, percent change from "
            "last year) and ED_Barrie_OneDay for the frequency tables (TriageLevel, Age or WaitTimeMin)."),
        "rules": (
            "- In this module, the frequency COUNTS themselves may be typed (they come from sorting and counting), "
            "but every total, relative, cumulative and percentage cell must be a formula. Check typed counts with "
            "query_data (group_by or bins). If they used COUNTIF/COUNTIFS, say kindly that it works but this week "
            "is about building the table by hand; it comes in a later week.\n"
            "- Absolute references are central this week. When a formula refers to a total or the HST rate, check "
            "for the $ and ask what would break without it if they dragged the formula down.\n"
            "- Watch the classic mistakes: relative frequencies that don't sum to 1 (or 100%), the last cumulative "
            "frequency not equal to the total, overlapping class limits (0-30, 30-60), percentages stored as 13 "
            "instead of 0.13 or 13%, and rounding with formatting vs ROUND."),
    },
    "Week 4 · Cleaning messy data": {
        "datasets": ["ED_Intake_Raw"],
        "focus": (
            "Cleaning a messy raw export in Excel: TRIM, PROPER, UPPER, LEFT, VALUE, TEXTBEFORE and TEXTAFTER "
            "(split BloodPressure 120/80 into Systolic and Diastolic, then VALUE), IF, AND, OR (e.g. a blood pressure "
            "category or a flag), XLOOKUP (site name from the site code via the Sites sheet), COUNTIF to count "
            "problems such as NA, and handling missing data (NA, n/a, blanks) so it doesn't break formulas "
            "(IF(G2=\"NA\",\"\",...), IFERROR, or checking with ISNUMBER). One request = one cleaning task with a "
            "clear check number at the end (e.g. 'how many readings are High?', 'average systolic?', 'how many "
            "visits were at Orillia?')."),
        "rules": (
            "- Attach ED_Intake_Raw (whole file) with the first request. Later requests keep working on the same "
            "file.\n"
            "- The analyst's workbook will be large. You'll see each formula column summarised (first formulas plus "
            "a summary of all results). Use the summary to check: compare their counts and averages with "
            "query_data on the cleaned dataset, e.g. count where BPCategory = High, mean of Systolic, count where "
            "Age is_blank.\n"
            "- Use the raw export facts below to check counting tasks (how many NA, how many spellings of Sex). "
            "If their count is off by the n/a or blank rows, hint that COUNTIF(...,\"NA\") misses n/a and blanks, "
            "and ask how they'd catch those.\n"
            "- Good habits to coach: keep the raw column and clean into a new column, never type over raw data; "
            "decide what NA should become (blank, not 0) and explain why 0 would distort an average."),
    },
    "Week 5 · Descriptive statistics and Tables": {
        "datasets": ["ED_Visits_2025"],
        "focus": (
            "Descriptive statistics in Excel: COUNT, AVERAGE, MEDIAN, MIN, MAX, range (MAX-MIN), VAR.S, STDEV.S, "
            "QUARTILE.INC for Q1 and Q3, and IQR (Q3-Q1). COUNTIFS to build frequency crosstabs (e.g. Site by "
            "Disposition, or AgeGroup by TriageLevel) with labelled rows and columns and mixed references. Excel "
            "Tables (Ctrl+T) and structured references like Table1[WaitTimeMin]. Requests need 2-4 numbers, "
            "sometimes split by one category, or one small crosstab."),
        "rules": (
            "- Check quartiles with q1/q3 (QUARTILE.INC). If their Q1/Q3 differ slightly, ask whether they used "
            "QUARTILE.EXC.\n"
            "- For crosstabs, check with query_data(metric='count', group_by=..., group_by2=...). Ask them to add "
            "row and column totals that match the grand total."),
    },
    "Week 6 · Random numbers and dates": {
        "datasets": ["ED_Visits_Dates"],
        "focus": (
            "Random numbers in Excel: RAND, RANDBETWEEN, RANDARRAY; drawing a random sample of visits for a chart "
            "audit (add a RAND() column and sort by it, or SORTBY/TAKE); why the numbers change on every "
            "recalculation and how to freeze them (Paste Special > Values). Dates in Excel: dates are serial "
            "numbers; TODAY, YEAR, MONTH, DAY, WEEKDAY, TEXT(date,\"ddd\") or \"mmm\", date arithmetic (days between "
            "two dates), DATEDIF for age in years, EDATE, NETWORKDAYS; and summarising by date parts (visits by "
            "weekday or month, with COUNTIFS on a helper column)."),
        "rules": (
            "- Random results can't be checked against fixed numbers. Check the METHOD instead: the formula, the "
            "range (e.g. RANDBETWEEN(1,300) gives whole numbers 1-300), sample size, no duplicate visits in a sample, "
            "and that they froze the values before reporting. Ask what would happen if they pressed F9.\n"
            "- Check date work with the check-only columns: AgeAtVisit, DaysToFollowUp, VisitWeekday/VisitWeekdayNum, "
            "VisitMonthNum, FollowUpBooked. If their ages are 1 too high for some rows, they probably did "
            "YEAR(VisitDate)-YEAR(DateOfBirth); hint at birthdays later in the year.\n"
            "- Watch for: days-between showing as a date (needs General/Number format), WEEKDAY numbering (1 = "
            "Sunday by default), blank FollowUpDate giving a huge negative number, and TODAY() making results "
            "change from day to day."),
    },
}
TOPICS = {k: v["focus"] for k, v in MODULES.items()}   # used by the sidebar and the check-in


def schema_text(topic: str) -> str:
    names = MODULES[topic]["datasets"]
    text = "Datasets for this module (fictional data):\n" + "\n".join(f"- {DATASET_DESC[n]}" for n in names)
    if "ED_Intake_Raw" in names:
        text += "\n" + raw_messy_facts()
    return text


def system_prompt(topic: str, student: str) -> str:
    return f"""You are Dana Morgan, Manager of Decision Support at Lakeshore Regional Health, a fictional hospital network in Central Ontario. You are chatting with {student}, a junior healthcare data analyst on your team. In reality they are a college student in a health informatics course, a beginner with Microsoft 365 Excel. Stay in character the whole time.

{schema_text(topic)}

This module's focus: {MODULES[topic]['focus']}

How you work:
- Write like a busy, supportive manager in a work chat: short, plain, friendly. Under 150 words per reply, except a new request may run to 200. No headings. Short bullet lists are fine.
- Each request gives: why you need it (business context), the specific question(s), who it's for, and when (e.g. "by end of day"). Ask for at most 3 things per request.
- Every new request (including the first) ends with one prediction question to answer BEFORE they start, e.g. "Before you start: do you expect Barrie's median wait to be above or below its average? Why?" Also call open_request with a short title and that question: the app shows them a prediction box and won't accept their results for the request until they've answered it. Their answer arrives as a message starting with "[Prediction]". Don't say whether it's right; just acknowledge it ("Noted, let's see what the data says") and let them get to work. When their results come in, compare them with the prediction and ask what surprised them, if anything.
- Stay within this module's Excel skills and a beginner's level. "A little harder" means a new question, an extra split or a judgement call using the same functions, never new techniques (no macros, no functions from a later module). Requests and their data come from Lakeshore's ED and its supplies.
- If you write your message and then call a tool, don't repeat that message after the tool result. Add only what's new, or nothing.
- Send data with attach_dataset, naming the dataset. Attach the dataset a request needs with your first request, and again whenever a request switches to another dataset. Later requests may come with a focused extract (one site, one quarter, fewer columns) when that fits. Every attachment goes in the same message as the full request it belongs to: never send a file without saying exactly what to do with it.
- One thing at a time. Don't stack a new request on top of an unfinished one. When you ask a follow-up question, wait for the answer, respond to it, and only then send the next request.
- Now and then leave one detail a little vague so the analyst has to ask a clarifying question. Answer clarifying questions the way a real manager would.
- Numbers count only when they come in the analyst's uploaded Excel workbook. If they type results without attaching the workbook, ask for the workbook before you confirm anything; you can still answer questions and clarify in the meantime. In the workbook, each result should be a formula (shown as "B4: =AVERAGE(...) -> 80.8"). A result typed in as a plain number ("B4: 80.8") doesn't count: ask them to replace it with the formula that produces it. Module-specific exceptions are listed below.
- Before you comment on ANY number the analyst reports, check it with the query_data tool (on the right dataset). Never guess. If they attached their workbook, review their formulas and results the same way and refer to specific cells (e.g. "your formula in B4"). Only call a cell right or wrong if you can actually see it. Remember an extract you sent has fewer rows than the full dataset, so check against the same filters.
- If a number is right, say so specifically. If it is off, say which figure doesn't match what you're seeing and give a hint (which function, range, filter or step to recheck). Don't reveal the correct value until they've missed it twice; then show it and explain briefly.
- The same goes for interpretation. If their takeaway is wrong, point them back at their own numbers with a guiding question ("compare B10 and B11: which group waits longer?") before you explain the right reading. Explain it only after a second miss.
- Occasionally (about every second or third request, and always when a result looks copied rather than built, e.g. correct numbers with formulas they don't seem to understand), once their results are right, run a quick check that they understand their own work: ask them to explain one of their formulas in their own words ("walk me through B5"), or to predict a small change ("if you switched B5 to Orillia, would it go up or down, and roughly to what?"). Check any number they give with query_data. Pick a different cell or change each time.
- Then ask one realistic follow-up a manager would ask: what explains it, why it matters, what we should do, or who else should see it. 
- Ask at most two questions per message. When you ask two, number them (1. and 2.) so neither gets lost. Never more than two.
- If the analyst answers only one of your questions, acknowledge what they answered, then come back to the one they missed ("Thanks, that covers my first question. What about the second: ...?"). Don't move on until both are answered.
- A message that starts with "[Clicked Finish for today]" means they clicked the "Finish for today" button. Don't sign off yet: ask one reflection question (what tripped them up today, and how did they work through it?). After they answer, respond in a sentence or two and sign off warmly. Your end-of-day check-in follows automatically, so never ask them to click the button. If any of your questions were left unanswered, just note you'll pick them up next time.
- Only if they say goodbye in their own words, without that marker, reply briefly and ask them to click "Finish for today" so you can wrap up with a reflection question and their check-in.
- Don't do the analysis for them. Show an example formula only if they're stuck after a hint.
- If they ask you to just give them the answers, or go off topic, redirect them kindly, in character.
- A request is finished when the prediction, results, any understanding check, and follow-up are done. Then thank them briefly and send the next request in the same message.

This module's checking rules:
{MODULES[topic]['rules']}

Excel facts to get right (Microsoft 365):
- There is no MEDIANIFS. For a median by group: =MEDIAN(FILTER(Table1[WaitTimeMin], Table1[Site]="Barrie")).
- TEXTBEFORE and TEXTAFTER return text, so wrap them in VALUE to get a number. TRIM removes leading, trailing and doubled spaces. PROPER capitalises the first letter of each word. COUNTIF criteria are not case-sensitive.
- XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found]); lock the lookup ranges with $ (or use a Table) before filling down.
- VAR.S and STDEV.S are for samples; QUARTILE.INC matches Excel's QUARTILE.
- DATEDIF(start, end, "Y") gives whole years; WEEKDAY(date) returns 1 for Sunday by default.
- Table column references (e.g. Barrie[TriageLevel]) shift to the next columns when a formula is dragged across with the fill handle. To lock one, write Barrie[[TriageLevel]:[TriageLevel]].
- Never guess the analyst's sheet names, table names, column letters or ranges. Use what you can see in their workbook, or ask.
- Never mention that you are an AI, a simulation or a tool."""


# ---------------------------------------------------------------- Claude call
def as_input_block(b) -> dict:
    """Copy a block of Claude's reply into the history with only the fields the API accepts back.
    (Newer SDK versions add extra response-only fields, which the API rejects if echoed.)"""
    if b.type == "text":
        return {"type": "text", "text": b.text}
    return {"type": "tool_use", "id": b.id, "name": b.name, "input": b.input}


def api_error_detail(e: Exception) -> str:
    """Short, readable reason from an API error, for the instructor."""
    msg = getattr(e, "message", "") or str(e)
    body = getattr(e, "body", None)
    if isinstance(body, dict) and isinstance(body.get("error"), dict):
        msg = body["error"].get("message", msg)
    return f"{e.__class__.__name__}: {msg}"[:300]


@st.cache_resource
def client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])


def ask_dana() -> tuple[str, list]:
    """Run the tool loop until Dana produces a text reply. Returns (text, attached files)."""
    model = st.secrets.get("MODEL", "claude-sonnet-5-5")
    system = [{"type": "text", "text": system_prompt(st.session_state.topic, st.session_state.student),
               "cache_control": {"type": "ephemeral"}}]
    msgs = st.session_state.api_messages
    st.session_state.outbox = []
    # Dana often writes part of her reply (e.g. the whole request) and calls a tool in the same step,
    # then adds a line after the tool result. Keep the text from every step, not just the last one.
    parts = []
    for _ in range(8):  # safety limit on tool rounds
        resp = client().messages.create(model=model, max_tokens=4000, system=system, tools=TOOLS, messages=msgs)
        u = resp.usage
        st.session_state.tokens["in"] += u.input_tokens + (u.cache_read_input_tokens or 0) + (u.cache_creation_input_tokens or 0)
        st.session_state.tokens["out"] += u.output_tokens
        msgs.append({"role": "assistant", "content": [as_input_block(b) for b in resp.content if b.type in ("text", "tool_use")]})
        step_text = "".join(b.text for b in resp.content if b.type == "text").strip()
        # Dana sometimes repeats her whole message after a tool call; keep only new text.
        repeated = any(step_text in p or difflib.SequenceMatcher(None, step_text, p).ratio() > 0.8
                       for p in parts)
        if step_text and not repeated:
            parts.append(step_text)
        if resp.stop_reason != "tool_use":
            if resp.stop_reason == "max_tokens" and parts:
                parts[-1] = trim_unfinished(parts[-1])
            return "\n\n".join(parts), st.session_state.outbox
        results = []
        for b in resp.content:
            if b.type == "tool_use":
                try:
                    out, err = json.dumps(TOOL_FUNCS[b.name](**b.input)), False
                except Exception as e:  # tell Dana what went wrong so she can retry
                    out, err = f"Error: {e}", True
                results.append({"type": "tool_result", "tool_use_id": b.id, "content": out, "is_error": err})
        msgs.append({"role": "user", "content": results})
    fallback = "Sorry, got pulled into a meeting. Can you resend that?"
    msgs.append({"role": "assistant", "content": fallback})
    return "\n\n".join(parts + [fallback]), st.session_state.outbox


CHECKIN_PROMPT = """You are Dana Morgan, Manager of Decision Support at Lakeshore Regional Health. You've just finished a working session with {student}, a junior data analyst on your team (in reality, a beginner Excel student in a health informatics course). This session's focus: {topic}.

Write {student} a short end-of-day check-in, in the style of a gentle, constructive workplace performance review. It is formative feedback only: no grades, scores or ratings.

Use exactly this format (Markdown):

**End-of-day check-in: {student}, {module}**

**What went well**
2-3 sentences naming specific strengths, with concrete examples from their actual work today (a formula, a check they ran, a good question, a solid takeaway).

**Where to grow**
- *Skill name.* At most two bullets. Frame each as a skill to build, not a mistake. Say briefly what you noticed and one concrete way to practise it.

**Focus for next time**
One sentence: the single most useful habit to work on.

End with one warm, sincere sentence.

Rules:
- Base everything on the conversation below. Never invent strengths or problems. If they did little work, say so kindly and keep it short.
- Be warm and encouraging but honest: real feedback, delivered kindly.
- Mention their reflection answer if it showed insight.
- Under 220 words. Plain, friendly language. Never mention that you are an AI or that this is a simulation.

The conversation:
{transcript}"""


def write_checkin() -> str:
    """Dana's end-of-day performance-review style feedback, from the whole conversation."""
    convo = "\n\n".join(
        f"{'Dana' if m['who'] == 'Dana' else st.session_state.student}: {m['text']}"
        + (f" [attached workbook: {m['upload']}]" if m["upload"] else "")
        + (" [attached chart screenshot]" if m["img"] else "")
        for m in st.session_state.chat if not m.get("checkin"))
    module = st.session_state.topic.split(" · ")[0]
    prompt = CHECKIN_PROMPT.format(student=st.session_state.student, topic=TOPICS[st.session_state.topic],
                                   module=module, transcript=convo)
    resp = client().messages.create(model=st.secrets.get("MODEL", "claude-sonnet-5-5"), max_tokens=3000,
                                    messages=[{"role": "user", "content": prompt}])
    st.session_state.tokens["in"] += resp.usage.input_tokens
    st.session_state.tokens["out"] += resp.usage.output_tokens
    text = "".join(b.text for b in resp.content if b.type == "text").strip()
    return trim_unfinished(text) if resp.stop_reason == "max_tokens" else text


def trim_unfinished(text: str) -> str:
    """If a reply ran out of room, cut it back to its last complete sentence."""
    cut = max(text.rfind(". "), text.rfind(".\n"), text.rfind("! "), text.rfind("? "))
    return text[:cut + 1] if cut > len(text) // 2 else text


# ---------------------------------------------------------------- keep conversations across reconnects
# Streamlit forgets a browser's session when it disconnects (a student switches to Excel for a while,
# a phone locks). Conversations are kept on the server under the student number, so a returning
# student only types the last four digits of their student number. A random link token in the page address also restores
# the conversation on a plain reload, without putting the student number in the address.
# Saved conversations are lost only when the app itself restarts (after a long idle sleep or a new deploy).
SAVED_KEYS = ["student", "student_number", "topic", "api_messages", "chat", "tokens", "wrap_stage", "request"]
KEEP_DAYS = 14


@st.cache_resource
def conversation_store() -> dict:
    return {"by_number": {}, "by_link": {}}


def norm_code(raw: str) -> str:
    """Class codes ignore capitals, spaces and dashes, so HIFN1009W27 = hifn1009-w27."""
    return re.sub(r"[\s\-_]+", "", raw or "").lower()


def clean_number(raw: str) -> str:
    return re.sub(r"\s+", "", raw or "")


def save_conversation():
    number = st.session_state.get("student_number")
    if not number:
        return
    store = conversation_store()
    store["by_number"][number] = {"saved": time.time(),
                                  **{k: st.session_state[k] for k in SAVED_KEYS if k in st.session_state}}
    cutoff = time.time() - KEEP_DAYS * 86400
    for old in [n for n, v in store["by_number"].items() if v["saved"] < cutoff]:
        store["by_number"].pop(old, None)
    for tok in [t for t, n in store["by_link"].items() if n not in store["by_number"]]:
        store["by_link"].pop(tok, None)


def link_session(number: str):
    """Put a random token in the page address so a reload restores this conversation."""
    token = secrets.token_urlsafe(8)
    conversation_store()["by_link"][token] = number
    st.query_params["s"] = token


def restore_conversation(number: str) -> bool:
    saved = conversation_store()["by_number"].get(number)
    if not saved:
        return False
    for k in SAVED_KEYS:
        st.session_state[k] = saved.get(k, 0 if k == "wrap_stage" else None)
    st.session_state.outbox = []
    return True


def reset(topic=None):
    st.session_state.wrap_stage = 0
    st.session_state.request = None
    st.session_state.api_messages = []
    st.session_state.chat = []  # dicts: who, text, img, files, upload
    st.session_state.tokens = {"in": 0, "out": 0}
    st.session_state.outbox = []
    if topic:
        st.session_state.topic = topic
    save_conversation()


if "student" not in st.session_state and st.query_params.get("s"):
    linked = conversation_store()["by_link"].get(st.query_params["s"])
    if not (linked and restore_conversation(linked)):
        st.query_params.clear()

# ---------------------------------------------------------------- sign-in gate
if "student" not in st.session_state:
    st.title("Analyst Inbox")
    st.write("Dana Morgan, Manager of Decision Support at Lakeshore Regional Health, will send you "
             "analysis requests. Do the work in Excel and reply with your results.")
    with st.form("signin"):
        number = st.text_input("Last 4 digits of your student number", max_chars=5)
        st.caption("Coming back to a conversation? These four digits are all you need.")
        name = st.text_input("First name (first visit only)")
        code = st.text_input("Class code (first visit only)", type="password")
        if st.form_submit_button("Open my inbox", type="primary"):
            number = clean_number(number)
            saved = conversation_store()["by_number"].get(number)
            other_name = saved and name.strip() and saved["student"].lower() != name.strip().lower()
            if not re.fullmatch(r"\d{4,5}", number):
                st.error("Enter the last 4 digits of your student number.")
            elif saved and not other_name:
                restore_conversation(number)
                link_session(number)
                st.rerun()
            elif not st.secrets.get("CLASS_CODE") or norm_code(code) != norm_code(str(st.secrets["CLASS_CODE"])):
                st.error("No saved conversation was found for those digits, so this is a new start. "
                         "Enter your first name and the class code your instructor posted on Blackboard.")
            elif not name.strip():
                st.error("Enter your first name so Dana knows who she's talking to.")
            elif other_name:
                st.error("Another student in this class has the same last four digits. Use the last "
                         "5 digits of your student number instead, here and every time you come back.")
            else:
                st.session_state.student = " ".join(w[:1].upper() + w[1:] for w in name.split())
                st.session_state.student_number = number
                reset(next(iter(TOPICS)))
                link_session(number)
                st.rerun()
    st.stop()

# ---------------------------------------------------------------- sidebar
if st.session_state.topic not in TOPICS:   # a session saved under an older module list
    reset(next(iter(TOPICS)))
with st.sidebar:
    st.subheader("This module's focus")
    topic = st.selectbox("Focus", list(TOPICS), index=list(TOPICS).index(st.session_state.topic),
                         label_visibility="collapsed")
    if topic != st.session_state.topic:
        reset(topic)
        st.rerun()

    st.subheader("Conversation")
    transcript = f"Student: {st.session_state.student} (student number ending {st.session_state.student_number})\n\n" + "\n\n".join(
        f"{'Dana Morgan' if m['who'] == 'Dana' else st.session_state.student}: {m['text']}"
        + "".join(f"\n[Attachment: {f['name']}]" for f in m["files"])
        + (f"\n[Attachment: {m['upload']}]" if m["upload"] else "")
        + ("\n[Attachment: chart screenshot]" if m["img"] else "")
        for m in st.session_state.chat)
    st.download_button("Download transcript", transcript,
                       f"transcript_{st.session_state.student}_{st.session_state.student_number}_{datetime.now():%Y%m%d}.txt",
                       width="stretch")
    if st.button("Finish for today", type="primary", width="stretch",
                 disabled=(not st.session_state.chat or st.session_state.get("wrap_stage", 0) == 1
                           or st.session_state.chat[-1].get("checkin", False))):
        st.session_state.wrap_up = True
    if st.button("Start over", width="stretch"):
        reset()
        st.rerun()
    st.caption("Click Finish for today and answer Dana's last question, then download your transcript "
               "and submit it on Blackboard.")

    st.caption(f"Signed in as {st.session_state.student} (…{st.session_state.student_number}). If the page "
               "reloads or you come back later, enter these digits to continue.")

    with st.expander("Lost Dana's file?"):
        for name in MODULES[st.session_state.topic]["datasets"]:
            st.download_button(f"Download {name}", dataset_file(name), f"{name}.xlsx", XLSX_MIME,
                               width="stretch", key=f"lost_{name}")

# ---------------------------------------------------------------- chat
st.markdown("#### Dana Morgan")
st.caption("Manager, Decision Support · Lakeshore Regional Health")

for i, m in enumerate(st.session_state.chat):
    with st.chat_message("assistant" if m["who"] == "Dana" else "user", avatar="🩺" if m["who"] == "Dana" else None):
        if m.get("checkin"):
            with st.container(border=True):
                st.markdown(m["text"])
        else:
            st.markdown(m["text"])
        if m["img"]:
            st.image(m["img"])
        if m["upload"]:
            st.caption(f"📎 {m['upload']}")
        for j, f in enumerate(m["files"]):
            st.download_button(f"📎 {f['name']}", f["data"], f["name"], f["mime"], key=f"dl_{i}_{j}")


def prediction_pending() -> bool:
    req = st.session_state.get("request")
    return bool(req) and not req.get("prediction")


def queue_prediction():
    text = st.session_state.get("prediction_box", "").strip()
    if len(text) < 10:
        st.session_state.prediction_warning = True
        return
    st.session_state.request["prediction"] = text
    st.session_state.pending_prediction = text
    st.session_state.prediction_box = ""


def prediction_box():
    req = st.session_state.request
    with st.container(border=True):
        st.markdown(f"**Before you start:** {req['question']}")
        st.text_area("Your prediction, and why", key="prediction_box", height=90,
                     placeholder="I expect ... because ...")
        if st.session_state.pop("prediction_warning", False):
            st.warning("Write a sentence or two: what you expect, and why.")
        st.button("Save my prediction", type="primary", on_click=queue_prediction)
        st.caption("You can still ask Dana questions in the chat. "
                   + ("Sending query results unlocks once you've saved your prediction."
                      if SQL_APP else "Attaching your workbook unlocks once you've saved your prediction."))


def post_dana(text: str, files: list):
    st.session_state.chat.append({"who": "Dana", "text": text, "img": None, "files": files, "upload": None})
    save_conversation()


if not st.session_state.chat:
    st.info(f"Welcome to the team, {st.session_state.student}. Dana's first request is on its way.")
    st.session_state.api_messages = [{"role": "user", "content": "Hi Dana, I'm ready for my first request."}]
    save_conversation()
    try:
        with st.spinner("Dana is typing…"):
            reply, files = ask_dana()
        if not files:  # make sure the analyst always gets the data with the first request
            first = MODULES[st.session_state.topic]["datasets"][0]
            files = [{"name": f"{first}.xlsx", "data": dataset_file(first), "mime": XLSX_MIME}]
        post_dana(reply, files)
        st.rerun()
    except anthropic.APIError as e:
        st.session_state.api_messages = []
        save_conversation()
        st.error("Dana can't be reached right now. Wait a minute and reload the page. "
                 f"(Details for your instructor: {api_error_detail(e)})")
        st.stop()

if st.session_state.pop("checkin_failed", False):
    st.warning("Dana couldn't write your end-of-day check-in just now. Click Finish for today to try again.")

sent = sum(1 for m in st.session_state.chat if m["who"] != "Dana")
if sent >= MAX_STUDENT_MESSAGES:
    st.warning("You've reached the message limit for this session. Download your transcript, then press Start over.")
    st.stop()

if prediction_pending():
    prediction_box()
    prompt = st.chat_input("Ask Dana a question. Save your prediction to unlock attaching your workbook.")
    prompt = {"text": prompt, "files": []} if prompt else None
else:
    prompt = st.chat_input("Reply to Dana. Attach your Excel workbook or a chart screenshot.",
                           accept_file=True, file_type=["xlsx", "png", "jpg", "jpeg"])
    prompt = {"text": prompt.text, "files": list(prompt.files)} if prompt else None
wrap_up = st.session_state.pop("wrap_up", False)
prediction = st.session_state.pop("pending_prediction", None)
if prompt or wrap_up or prediction:
    if prediction:
        text = f"My prediction: {prediction}"
    else:
        text = (prompt["text"] or "").strip() if prompt else "I need to stop for today."
    entry = {"who": "You", "text": text, "img": None, "files": [], "upload": None}
    content = []
    for f in (prompt["files"][:1] if prompt else []):
        data, lower = f.getvalue(), f.name.lower()
        if lower.endswith(".xlsx"):
            entry["upload"] = f.name
            content.append({"type": "text", "text": read_workbook(data, f.name)})
        else:
            entry["img"] = data
            media = "image/png" if lower.endswith(".png") else "image/jpeg"
            content.append({"type": "image", "source": {"type": "base64", "media_type": media,
                                                        "data": base64.b64encode(data).decode()}})
    if not text:
        text = entry["text"] = "(See attached.)"
    api_text = (f"[Prediction] {prediction}" if prediction
                else f"[Clicked Finish for today] {text}" if wrap_up else text)
    content.append({"type": "text", "text": api_text})
    st.session_state.chat.append(entry)
    answering_reflection = st.session_state.get("wrap_stage", 0) == 1 and not wrap_up
    before = len(st.session_state.api_messages)
    st.session_state.api_messages.append({"role": "user", "content": content})
    try:
        with st.spinner("Dana is reviewing…" if entry["upload"] or entry["img"] else "Dana is typing…"):
            reply, files = ask_dana()
        post_dana(reply, files)
    except anthropic.APIError as e:
        st.session_state.chat.pop()
        del st.session_state.api_messages[before:]  # drop the message and any half-finished tool rounds
        st.error("Your message didn't go through. Wait a moment and send it again. "
                 f"(Details for your instructor: {api_error_detail(e)})")
        st.stop()
    if wrap_up:
        st.session_state.wrap_stage = 1          # Dana has asked the reflection question
    elif answering_reflection:
        st.session_state.wrap_stage = 0
        try:
            with st.spinner("Dana is writing your end-of-day check-in…"):
                note = write_checkin()
            st.session_state.chat.append({"who": "Dana", "text": note, "img": None, "files": [],
                                          "upload": None, "checkin": True})
        except anthropic.APIError:
            st.session_state.checkin_failed = True
    save_conversation()
    st.rerun()
