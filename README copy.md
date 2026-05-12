# CSD3 Weekly Report Processor

Streamlit version of the CSD3 weekly reports notebook.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app opens at `http://localhost:8501`.

## Inputs

1. **Students workbook** — `Students_CSD3_Weekly Reports...xlsx`
2. **Adults workbook** — `Adults_CSD3_Weekly Reports...xlsx`
3. **All workbook** — `All_CSD3_Weekly Reports...xlsx`
4. **Target values** — one integer per line, in the same site order as the source data (e.g. `152`, `200`, `100`).
5. **Output filename** — name for the downloaded `.xlsx` (extension added automatically).

## Output

A single Excel file with these sheets:

- Student Summary Statistics (with red/green % coloring)
- Missing Student Summary
- Pull out - Missing Site Info
- Missing - {site} (one tab per school)
- Pull out - Young DOB
- Missing Staff Summary
- Pull out - Missing Staff Info
- One sheet per site with activity-session details

Red highlights flag missing/blank values; blue highlights flag suspect DOBs.

## Deploy (optional)

Push `app.py` and `requirements.txt` to a GitHub repo and deploy free at [share.streamlit.io](https://share.streamlit.io).
