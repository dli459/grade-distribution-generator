import io
import re
import zipfile
from pathlib import Path

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

GRADES = ['A+', 'A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-']
COLORS = ['#24638f', '#408bbb', '#8bbbd9', '#237b55', '#4aaa7d', '#91cbaa',
          '#c96b19', '#e99b4c', '#f4c08b', '#bc3038', '#db666b', '#e9a2a5']
X = np.array([0, 1, 2, 4.2, 5.2, 6.2, 8.4, 9.4, 10.4, 12.6, 13.6, 14.6])

st.set_page_config(page_title='Grade Distribution Generator', page_icon='📊', layout='wide')
st.title('Grade Distribution Generator')
st.caption('Upload an Excel workbook → preview courses → export Headcount and Percentage charts')

with st.expander('Expected Excel format', expanded=False):
    st.write('Use a worksheet named **Raw_Data**, or put the data on the first worksheet. Required columns: **Course, Year, Term** and the 12 grade columns from **A+** through **D-**. Optional: **Subject** and **Course Headcount**.')
    st.write('One row per course/semester. Blank or non-numeric grade cells are treated as missing, not zero; rows with no grades are skipped. Percentages use the sum of the 12 graded counts, excluding unknown/missing grades.')

file = st.file_uploader('Choose an Excel file', type=['xlsx'])
if not file:
    st.info('Upload your Excel file to begin. No account or API key is required.')
    st.stop()

try:
    book = pd.ExcelFile(file)
    sheet = 'Raw_Data' if 'Raw_Data' in book.sheet_names else book.sheet_names[0]
    df = pd.read_excel(book, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    required = ['Course', 'Year', 'Term'] + GRADES
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error('Missing columns: ' + ', '.join(missing))
        st.stop()
    for grade in GRADES:
        df[grade] = pd.to_numeric(df[grade], errors='coerce')
    invalid = df[GRADES].lt(0).any(axis=1) | df[GRADES].apply(lambda s: (s.notna() & (s % 1 != 0))).any(axis=1)
    if invalid.any():
        st.error(f'{invalid.sum()} row(s) contain negative or non-integer grade counts. Please correct the source file.')
        st.dataframe(df.loc[invalid, ['Course', 'Year', 'Term'] + GRADES], hide_index=True)
        st.stop()
    df['Graded Total'] = df[GRADES].sum(axis=1, min_count=1)
    skipped = df['Graded Total'].isna() | df['Graded Total'].eq(0)
    df = df.loc[~skipped].copy()
    if df.empty:
        st.error('No rows with graded students were found.')
        st.stop()
    def label(row):
        year = str(row['Year']).replace('.0', '')
        subject = str(row.get('Subject', '')).strip() if pd.notna(row.get('Subject', '')) else ''
        course = str(row['Course']).replace('.0', '').strip()
        course_name = f'{subject} {course}'.strip()
        return f"{course_name} | {year} {str(row['Term']).strip()}"
    df['Selection'] = df.apply(label, axis=1)
    df['Selection'] = df['Selection'] + df.groupby('Selection').cumcount().map(lambda n: f' (row {n+1})' if n else '')
except Exception as exc:
    st.error(f'Could not read this workbook: {exc}')
    st.stop()

st.success(f'Loaded {len(df)} course rows from “{sheet}”. Skipped {int(skipped.sum())} rows with no graded students.')
st.dataframe(df[['Selection', 'Graded Total'] + GRADES], hide_index=True, use_container_width=True)
mode = st.radio('Generate charts for', ['Selected course', 'All courses'], horizontal=True)
selected = st.selectbox('Course / Semester', df['Selection'].tolist(), disabled=(mode == 'All courses'))


def draw_chart(row, kind):
    counts = row[GRADES].to_numpy(dtype=float)
    total = counts.sum()
    values = counts if kind == 'Headcount' else counts / total * 100
    fig, ax = plt.subplots(figsize=(12.5, 5.5), dpi=160)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    bars = ax.bar(X, values, width=0.94, color=COLORS, edgecolor='none')
    ax.set_xticks(X, GRADES)
    ax.set_xlim(-0.65, 15.25)
    ax.set_ylabel('Count' if kind == 'Headcount' else 'Percentage (%)')
    ax.set_title(f"{row['Selection']} — {kind}", pad=17, fontsize=13, weight='bold')
    ax.grid(axis='y', linestyle='--', alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines[['top', 'right']].set_visible(False)
    ceiling = max(values.max() * 1.20, 1)
    ax.set_ylim(0, ceiling)
    for bar, val in zip(bars, values):
        label = f'{int(round(val))}' if kind == 'Headcount' else f'{val:.2f}%'
        ax.annotate(label, (bar.get_x() + bar.get_width()/2, bar.get_height()),
                    xytext=(0, 4), textcoords='offset points', ha='center',
                    va='bottom', fontsize=8, color='#243242', fontweight='semibold')
    fig.tight_layout()
    output = io.BytesIO()
    fig.savefig(output, format='png', bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return output.getvalue()


def safe_name(value):
    return re.sub(r'[^A-Za-z0-9._-]+', '_', value).strip('_')[:100]

if mode == 'Selected course':
    row = df.loc[df['Selection'].eq(selected)].iloc[0]
    count_png = draw_chart(row, 'Headcount')
    pct_png = draw_chart(row, 'Percentage')
    left, right = st.columns(2)
    with left:
        st.image(count_png, caption='Headcount')
        st.download_button('Download Headcount PNG', count_png, f'{safe_name(selected)}_Headcount.png', 'image/png')
    with right:
        st.image(pct_png, caption='Percentage')
        st.download_button('Download Percentage PNG', pct_png, f'{safe_name(selected)}_Percentage.png', 'image/png')
    rows_to_export = [row]
else:
    rows_to_export = [row for _, row in df.iterrows()]
    st.write(f'Ready to export **{len(rows_to_export) * 2} PNG charts**.')

zip_buffer = io.BytesIO()
with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
    for row in rows_to_export:
        name = safe_name(row['Selection'])
        archive.writestr(f'{name}_Headcount.png', draw_chart(row, 'Headcount'))
        archive.writestr(f'{name}_Percentage.png', draw_chart(row, 'Percentage'))
st.download_button('📦 Download all charts as ZIP', zip_buffer.getvalue(), 'Grade_Distribution_Charts.zip', 'application/zip', type='primary')

st.caption('The uploaded workbook is processed in the app session; this prototype does not save it to a database. Hosting providers may still process uploaded data, so check your institution’s data policy before uploading private records.')
