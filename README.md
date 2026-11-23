# Grade Distribution Generator

A lightweight web app for turning university course grade-count spreadsheets into consistent, publication-ready charts. Built with **Python, Streamlit, pandas and Matplotlib**. Users can upload an Excel file, choose courses, preview results, and download PNG images without installing Python.

> **Live app:** Add your Streamlit URL here after deployment.

## What the app can do

| Feature | Description |
| --- | --- |
| Single-course charts | Generate **Headcount** and **Percentage** charts for the 12 grades from A+ through D−. |
| Bulk export | Click **Generate ZIP** to build both charts for every eligible course record; follow the live progress bar, then click **Download all charts as ZIP**. |
| Course/year comparison | Select multiple course/year/semester records and generate a **100% stacked percentage chart** with A, B, C and D grade groups. |
| Downloadable input template | Download a blank Excel template and enter new course data. |
| Original Dashboard exports | Read supported *Student Grade Distributions Dashboard* Excel exports directly, without manually removing their title rows or empty columns. |
| Validation | Flag invalid grade counts and skip records with no A+–D− grades. |

## Quick start for users

1. Open the deployed Streamlit web app.
2. If you are entering data yourself, click **Download blank Excel template**, fill it in, and save it as `.xlsx`. Alternatively, use a supported original Dashboard export.
3. Upload your `.xlsx` file using **Choose an Excel file**.
4. Review the detected records and any skipped-row message.
5. For individual charts, select **Selected course** and choose the course/semester; or select **All courses**, click **Generate ZIP**, watch the progress indicator, and download the finished ZIP.
6. For a comparison chart, scroll to **Compare Courses / Years**, choose Course, Year and Semester filters, click **Apply filters**, and select the specific records to include. Download the PNG.

## Supported Excel formats

### Option A — original Grade Distributions Dashboard export

The app looks through worksheets for the expected header row in the **first 40 rows**. The recognized header labels are:

- `Year`, `Term Group`, `Subject`, `Catalog Number`
- `A+`, `A`, `A-`, `B+`, `B`, `B-`, `C+`, `C`, `C-`, `D+`, `D`, `D-`

The app combines `Subject` and `Catalog Number` into the displayed course name. Additional columns such as `Unknown`, `Withdrawn`, or `Not Graded` are not included in the percentage denominator.

### Option B — prepared input template

Download the template from the app. Its worksheet is named `Raw_Data`, with this header row:

```text
Course | Year | Term | A+ | A | A- | B+ | B | B- | C+ | C | C- | D+ | D | D-
```

One row represents one course in one year and term. For example, `COMPSYS 201`, `2026`, `Semester 1`. The example names are illustrative; enter your own counts. Keep the column headings unchanged. You can add as many rows as you need.

### Data requirements

- File type: `.xlsx`.
- `Year`: a four-digit year (for example, `2025`).
- `Course` and `Term`: non-empty.
- Grade counts: **non-negative whole numbers**. Blank grade cells in otherwise valid records are treated as zero.
- Rows where all 12 grade counts are missing or zero are skipped.
- The app may display separate bars for different semesters of the same course/year; it does **not** automatically combine semesters.
- For accurate comparisons, ensure the input represents distinct records; duplicate course/year/term records are not automatically merged.

## Chart rules

**Single-course charts:** the 12 grade levels use four related color families: A = blue, B = green, C = orange, D = red. Bars within a family are close together, with visible gaps between families. Headcount labels are whole numbers; percentage labels show **two decimal places**.

**Percentages:** the denominator is the sum of grade counts from A+ through D−, **not** a separate course headcount field. For example:

```text
A percentage = (A+ + A + A-) / (sum of A+ through D-) × 100
```

**Comparison filters:** use **Course**, **Year**, and **Semester / Term** independently. An empty filter means “all”. The final **Courses / semesters to include in the chart** control selects the individual records. Click **Apply filters** to update the available records. Form selections do not trigger a rerun until applied; selections that no longer match are removed. A “No results” message in an open dropdown can also mean there are no further options to add. Use **Select all matching records** to include every record matching the filters. The same course in different years or semesters remains separate (not automatically aggregated).

**Comparison chart:** A/B/C/D each combine their `+`, plain, and `−` grades. Each selected course/year/term record becomes one 100%-stacked bar. Percentages are calculated separately for each record. Small segments may omit labels to avoid clutter.

## Performance and progress

- **Course / Year / Semester** filter changes are grouped in an **Apply filters** form to avoid unnecessary reruns for each selection.
- Single-course and comparison PNGs are cached by their data and selection, so unrelated interactions do not redraw identical charts.
- Bulk ZIP generation starts **only** when **Generate ZIP** is clicked, and reports progress for each completed PNG. The ZIP download appears after generation completes.
- Streamlit may briefly dim the page while a script is running. This is expected during bulk generation, but filter changes should be much faster.
- The comparison preview still updates when the selected records change; very large selections can take longer the first time.

## Deploy to Streamlit Community Cloud

1. Put `app.py`, `requirements.txt`, and `README.md` in the root of your GitHub repository.
2. Sign in at https://share.streamlit.io/ and choose **Create app**.
3. Select the repository, its branch (normally `main`), and main file path `app.py`.
4. Deploy, then open the provided `.streamlit.app` URL.
5. Upload a **non-sensitive test workbook** and check the results before sharing the link.

### Update an existing deployment

Edit or replace `app.py` and `README.md` in the **same** GitHub repository and commit the changes. Streamlit Community Cloud normally redeploys when the configured branch changes. You do not need to create a second app. If it does not refresh, check the app's deployment logs and restart/reboot it from Streamlit's app management interface.

## Run locally (optional)

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Project files

```text
app.py            Streamlit interface, Excel parsing, chart generation, downloads
requirements.txt  Python package requirements
README.md         User guide and deployment instructions
```

## Privacy and limitations

**Important:** A public GitHub repository exposes its source code, and a publicly accessible Streamlit app may be used by anyone with its URL. Do not commit student-level records, confidential spreadsheets, API keys, or credentials. Only upload university data to a hosting service when your institution's data-handling policies permit it. The app does not intentionally write uploaded workbooks to a persistent database, but the hosting platform processes the files in its environment.

This is a prototype, not an official University of Auckland service. Input formats outside the two documented layouts may need adjustments. Charts should be checked against the source data before use in formal reports.
