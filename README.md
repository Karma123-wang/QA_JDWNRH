# BHSQA KPI Dashboard (Streamlit)

An interactive dashboard for presenting hospital KPI performance against the national
**BHSQA KPI User Guide** targets (QASD, Ministry of Health).

* Upload the hospital's KPI reporting workbooks (one per fiscal year, July–June – e.g. the
  2025-26 and 2026-27 QAS sheets). The app detects each file's fiscal year automatically.
* Only KPIs that align with the national BHSQA list are extracted.
* Pick any reporting period (e.g. Jan–Oct 2026) – months are stitched across the two files.
* **Tap any KPI** to see its trend, its monthly numerator/denominator and the **exact source
  cells** (file › sheet › cell), with the original sheet shown and the cells highlighted.
* A **national scorecard** lists every BHSQA KPI number with its description, target, hospital value, status and each unit's value.
* Download the result as **Excel** (with formulas, a Source Map and clickable links),
  **PDF** report or **PowerPoint** deck for senior officials.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit dashboard |
| `kpi_engine.py` | Reads the workbooks, maps units/KPIs to BHSQA, targets, status, source tracing |
| `exporters.py` | Excel / PDF / PowerPoint reports |
| `requirements.txt` | Python packages for Streamlit Community Cloud |
| `.streamlit/config.toml` | Colours and upload size |
| `data/` (optional) | Put the workbooks here if you want the app to open with data already loaded |

## Deploy on Streamlit Community Cloud (streamlit.io)

1. Create a new repository on GitHub (choose **Private** – the data is hospital data).
2. Upload all files from this folder, keeping the `.streamlit` folder
   (on github.com: **Add file → Upload files**, drag the whole folder in, **Commit**).
3. Go to <https://share.streamlit.io> → **Create app** → **Deploy a public app from GitHub**
   → choose your repository, branch `main`, main file `app.py`.
4. Under **Advanced settings** choose Python **3.12**, then **Deploy**.
5. For a private repo, use the app's **Share** button to invite the officials' e-mail addresses.

To present without uploading each time, create a `data` folder in the repo and put the
workbooks in it – the app loads them automatically when nothing is uploaded.
Only do this in a **private** repository.

## Run on your own computer

```bash
pip install -r requirements.txt
streamlit run app.py
```

## How the numbers are worked out

* **Achievement** for rate KPIs = total numerator ÷ total denominator over the months in the
  period that have *both* values (× 100, × 1000 or × 100,000 as the BHSQA formula requires).
  Survey KPIs (satisfaction, 5S, waiting time) = average of the survey rounds in the period.
* **Status**: *Met* / *Not met* against the BHSQA target; *No data* when no month has both
  values; *Not rated* when the BHSQA target cannot be applied to the data as recorded
  (e.g. needle-stick injuries recorded per OPD visit instead of per staff working day).
* Special handling (shown in the app's notes):
  * Bed occupancy uses bed strength × days in the month.
  * Discharge-summary and ICD-coding KPIs use discharges + deaths as the denominator.
  * Falls, bedsores and IPD blood/body-fluid exposure use hospital inpatient days.
  * SSI uses surgeries performed (from Medical Record).
  * Blood component usage follows the BHSQA formula: component units ÷ all units transfused.
* Times recorded as `hh:mm:ss` are converted to minutes; fractions such as `0.85` in survey
  cells are read as 85 %.
* Data-quality problems (text in number cells, numerators without denominators, values over
  100 %, zero denominators) are listed in the **Data quality** tab.

## Updating for a new year

Upload the new fiscal-year workbook alongside the previous one. Sheet names that changed
between years (e.g. `AICU` → `AICU(MICU I & II)`, `ICWM` → `IPC+AMSU`) are mapped in
`UNITS` in `kpi_engine.py`. Add an alias there if a unit sheet is renamed.
KPI rows are found by the KPI name in column B, so moving rows up or down doesn't matter.
