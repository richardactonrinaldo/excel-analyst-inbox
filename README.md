# Analyst Inbox

Students in HIFN1009 chat with **Dana Morgan**, a pretend Manager of Decision Support at a fictional Ontario hospital network. Dana sends analysis requests about fictional emergency department and supply data, checks students' numbers against the data, gives hints instead of answers, and asks follow-up questions.

Dana sends the data as an Excel attachment (the week's dataset with her first request, focused extracts later). Students can attach their own Excel workbook, and Dana reviews their formulas and results cell by cell, or a chart screenshot.

Module options (students pick one in the sidebar):

| Module | Skills | Dataset(s) |
| --- | --- | --- |
| Weeks 2–3 · Arithmetic and frequency tables | + − × ÷, percentages, number formats, relative/absolute references; frequency, relative, cumulative and grouped frequency tables built without COUNTIF | `Supply_Order` (10 rows), `ED_Barrie_OneDay` (30 rows) |
| Week 4 · Cleaning messy data | TRIM, PROPER, UPPER, LEFT, VALUE, TEXTBEFORE/TEXTAFTER (blood pressure 120/80), IF, AND, OR, XLOOKUP, COUNTIF, NA/blank handling | `ED_Intake_Raw` (60-row raw export + Sites lookup sheet) |
| Week 5 · Descriptive statistics and Tables | COUNT, AVERAGE, MEDIAN, MIN, MAX, range, VAR.S, STDEV.S, QUARTILE.INC, IQR, COUNTIFS crosstabs, Excel Tables | `ED_Visits_2025` (300 rows) |
| Week 6 · Random numbers and dates | RAND, RANDBETWEEN, RANDARRAY, random samples; INT(YEARFRAC(...,1)), TEXT(date,"dddd" / "yyyy-mm" / "yyyy-mmm" / "yyyy-mmmm"), WEEKDAY(date,2), date + days, YEAR, TODAY | `ED_Visits_Dates` (300 rows) |

Dana checks every number against a correct version of each dataset. For the messy Week 4 data she checks against a cleaned answer key (`data/ED_Intake_Clean.csv`), verified by recalculating a full set of cleaning formulas in a spreadsheet.

## Deploy on Streamlit Community Cloud (free)

1. **Get an API key.** Sign in at platform.claude.com, buy a small amount of credit (e.g. $20 USD), leave auto-reload **off** so that amount is a hard cap, and create an API key.
2. **Deploy.** Sign in at share.streamlit.io with GitHub, click **Create app**, choose this repo, branch `main`, main file `app.py`.
3. **Add secrets.** In the app's **Settings > Secrets**, paste:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   CLASS_CODE = "your-class-code"
   MODEL = "claude-sonnet-5-5"
   ```
4. **Share.** Post the app URL and the class code on Blackboard. Change `CLASS_CODE` in Secrets each term (or any time it leaks).

Usage and cost appear in the Claude Console under Usage.

## Cost controls

- The class code keeps strangers from using your credits.
- Each session is capped at 60 student messages (`MAX_STUDENT_MESSAGES` in `app.py`).
- Prepaid credits with auto-reload off can't be exceeded.
- `MODEL = "claude-haiku-5-5"` cuts cost by roughly 20x.

## Run locally

```bash
pip install -r requirements.txt
cp .streamlit/secrets.example.toml .streamlit/secrets.toml   # then edit it
streamlit run app.py
```

## Change the data

`python make_dataset.py` regenerates `data/ED_Visits_2025.csv`, and `python make_module_data.py` regenerates the Weeks 2–3, 4 and 6 files (fixed seeds, same result every time). Datasets are described in `DATASET_DESC` and wired up in `load_datasets()` and `MODULES` in `app.py`.
