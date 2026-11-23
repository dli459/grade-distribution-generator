# Grade Distribution Generator

A simple Streamlit app for generating two charts per course: **Headcount** and **Percentage**.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publish a website for Kevin

1. Create a GitHub repository and upload `app.py` and `requirements.txt` (and optionally this README).
2. Visit https://share.streamlit.io/ and sign in with GitHub.
3. Choose **Create app**, select the repository, branch `main`, and entry point `app.py`.
4. Deploy and share the resulting URL with Kevin.

**Privacy:** Do not upload private or restricted university datasets to public hosting without appropriate authorization. A public Streamlit URL is accessible to anyone who knows it unless additional access controls are configured. Use a private deployment if necessary.

## Input workbook

The app looks for a sheet named `Raw_Data`, otherwise uses the first sheet. Required columns: `Course`, `Year`, `Term`, `A+`, `A`, `A-`, `B+`, `B`, `B-`, `C+`, `C`, `C-`, `D+`, `D`, `D-`.

Optional columns include `Subject` and `Course Headcount`. Rows without graded counts are skipped. The denominator for percentages is the sum of A+ through D- counts, not `Course Headcount`.

## Output

- Per-course headcount PNG with integer labels
- Per-course percentage PNG with two-decimal labels
- Downloadable ZIP for selected or all courses
- A/B/C/D grade groups use blue/green/orange/red shades, with tightly spaced bars within each group and clear gaps between groups.
