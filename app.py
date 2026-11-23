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

with st.expander('Supported Excel formats', expanded=False):
    st.write('**Dashboard export:** the original Grade Distributions Dashboard workbook, including its title rows and split grade columns.')
    st.write('**Prepared data:** a Raw_Data sheet with Course, Year, Term, and A+ through D- columns.')
    st.write('Percentages are calculated using only A+ through D- grade counts. Rows with no graded students are skipped.')

file = st.file_uploader('Choose an Excel file', type=['xlsx'])
if not file:
    st.info('Upload your Excel file to begin. No account or API key is required.')
    st.stop()


def parse_workbook(upload):
    book = pd.ExcelFile(upload)
    for sheet in (['Raw_Data'] if 'Raw_Data' in book.sheet_names else []) + [x for x in book.sheet_names if x != 'Raw_Data']:
        preview = pd.read_excel(book, sheet_name=sheet, header=None, nrows=40, dtype=object)
        header_index = None
        is_dashboard = False
        for i, row in preview.iterrows():
            entries = {str(x).strip() for x in row.dropna()}
            if {'Year', 'Term Group', 'Subject', 'Catalog Number', 'A+', 'D-'}.issubset(entries):
                header_index = i
                is_dashboard = True
                break
            if {'Course', 'Year', 'Term', 'A+', 'D-'}.issubset(entries):
                header_index = i
                break
        if header_index is None:
            continue
        data = pd.read_excel(book, sheet_name=sheet, header=header_index, dtype=object)
        data.columns = [str(x).strip() for x in data.columns]
        if is_dashboard:
            data = data.rename(columns={'Term Group': 'Term', 'Catalog Number': 'Course'})
        required = ['Course', 'Year', 'Term'] + GRADES
        if any(c not in data.columns for c in required):
            continue
        data = data.dropna(subset=['Course', 'Year', 'Term'], how='any').copy()
        data = data[data['Course'].astype(str).str.strip().ne('')]
        data = data[data['Year'].astype(str).str.match(r'^\d{4}(?:\.0)?$')]
        if data.empty:
            continue
        for grade in GRADES:
            original = data[grade]
            converted = pd.to_numeric(original, errors='coerce')
            unexpected = original.notna() & original.astype(str).str.strip().ne('') & converted.isna()
            if unexpected.any():
                raise ValueError(f'Non-numeric grade counts found in column {grade}.')
            data[grade] = converted
        invalid = data[GRADES].lt(0).any(axis=1) | data[GRADES].apply(
            lambda col: (col.notna() & (col % 1 != 0))
        ).any(axis=1)
        if invalid.any():
            raise ValueError(f'{int(invalid.sum())} row(s) contain negative or non-integer grade counts.')
        data['Graded Total'] = data[GRADES].sum(axis=1, min_count=1)
        skipped = data['Graded Total'].isna() | data['Graded Total'].eq(0)
        data = data.loc[~skipped].copy()
        if data.empty:
            raise ValueError('No rows with graded students were found.')
        data[GRADES] = data[GRADES].fillna(0)
        def label(row):
            year = str(row['Year']).replace('.0', '')
            subject = str(row.get('Subject', '')).strip() if pd.notna(row.get('Subject', '')) else ''
            course = str(row['Course']).replace('.0', '').strip()
            course_name = course if course.upper().startswith(subject.upper() + ' ') and subject else f'{subject} {course}'.strip()
            return f"{course_name} | {year} {str(row['Term']).strip()}"
        data['Selection'] = data.apply(label, axis=1)
        data['Selection'] += data.groupby('Selection').cumcount().map(
            lambda n: f' (duplicate {n+1})' if n else ''
        )
        return data, sheet, int(skipped.sum()), 'Dashboard export' if is_dashboard else 'Prepared Raw_Data'
    raise ValueError('No supported grade table found. Check the workbook format and grade headers.')


try:
    df, sheet, skipped_count, format_name = parse_workbook(file)
except Exception as exc:
    st.error(f'Could not read this workbook: {exc}')
    st.stop()

st.success(f'Loaded {len(df)} course rows from “{sheet}”. Skipped {skipped_count} rows with no graded students. Detected format: {format_name}.')
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


# Cross-course / cross-year A-B-C-D comparison (one bar per selected row).
st.divider()
st.header('Compare Courses / Years — 100% Stacked Percentage')
st.write('Select the course/semester records you want in one comparison chart. Each bar sums to 100%.')

course_names = df['Selection'].str.split(' | ', regex=False).str[0]
course_options = sorted(course_names.unique().tolist())
preferred_course = course_options[0] if course_options else None
course_filter = st.multiselect(
    'Filter by course (optional)', course_options,
    default=[preferred_course] if preferred_course else [],
    help='Select one course to compare years, or several courses to compare them together.',
)
comparison_df = df.loc[course_names.isin(course_filter)] if course_filter else df
options = comparison_df['Selection'].tolist()
default_options = options[:min(len(options), 10)]
comparison_selections = st.multiselect(
    'Courses / semesters to include in the chart', options,
    default=default_options,
    help='Every selected record becomes one bar. You can remove individual years or semesters.',
)

STACK_LABELS = ['A Grades', 'B Grades', 'C Grades', 'D Grades']
STACK_COLORS = ['#22547c', '#315c9d', '#578638', '#c90000']

def draw_stacked_comparison(rows):
    rows = rows.copy()
    grade_counts = rows[GRADES].to_numpy(dtype=float)
    grouped_counts = np.column_stack([
        grade_counts[:, 0:3].sum(axis=1),
        grade_counts[:, 3:6].sum(axis=1),
        grade_counts[:, 6:9].sum(axis=1),
        grade_counts[:, 9:12].sum(axis=1),
    ])
    totals = grouped_counts.sum(axis=1)
    percentages = grouped_counts / totals[:, None] * 100.0
    n = len(rows)
    fig, ax = plt.subplots(figsize=(max(10, n * 1.18 + 2.5), 6.8), dpi=160)
    fig.patch.set_facecolor('white')
    x = np.arange(n)
    bottoms = np.zeros(n)
    for j, (label, color) in enumerate(zip(STACK_LABELS, STACK_COLORS)):
        heights = percentages[:, j]
        ax.bar(x, heights, bottom=bottoms, width=0.76, label=label,
               color=color, edgecolor='white', linewidth=0.7)
        for i, h in enumerate(heights):
            if h >= 4.0:
                ax.text(x[i], bottoms[i] + h / 2, f'{h:.1f}%',
                        ha='center', va='center', color='white',
                        fontsize=9, fontweight='bold')
        bottoms += heights
    selected_course_names = rows['Selection'].str.split(' | ', regex=False).str[0]
    if selected_course_names.nunique() == 1:
        ticks = [f"{str(y).replace('.0', '')} {t}" for y, t in zip(rows['Year'], rows['Term'])]
        title = f"Percentage Grade Breakdown: {selected_course_names.iloc[0]}"
    else:
        ticks = [str(x).replace(' | ', '\n') for x in rows['Selection']]
        title = 'Percentage Grade Breakdown: Selected Courses / Years'
    ax.set_xticks(x, ticks, rotation=35 if n > 8 else 0,
                  ha='right' if n > 8 else 'center', fontsize=9)
    ax.set_ylabel('Percentage (%)', fontweight='bold')
    ax.set_xlabel('Year / Semester' if selected_course_names.nunique() == 1 else 'Course / Year / Semester')
    ax.set_title(title, fontweight='bold', fontsize=14, pad=16)
    ax.set_ylim(0, 100)
    ax.set_yticks(np.arange(0, 101, 20))
    ax.grid(axis='y', color='#d5d5d5', linewidth=0.8)
    ax.set_axisbelow(True)
    ax.spines[['top', 'right']].set_visible(False)
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.19 if n <= 8 else -0.29),
              ncol=4, frameon=False, fontsize=9)
    fig.tight_layout()
    result = io.BytesIO()
    fig.savefig(result, format='png', bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return result.getvalue()

if comparison_selections:
    comparison_rows = df.set_index('Selection').loc[comparison_selections].reset_index()
    comparison_png = draw_stacked_comparison(comparison_rows)
    st.image(comparison_png, caption=f'{len(comparison_selections)} selected course/semester records')
    st.download_button('Download comparison PNG', comparison_png,
                       'Grade_Distribution_Comparison_Stacked.png', 'image/png',
                       key='comparison_download')
else:
    st.info('Choose at least one course/semester record to create the comparison chart.')

st.caption('The uploaded workbook is processed in the app session; this prototype does not save it to a database. Hosting providers may still process uploaded data, so check your institution’s data policy before uploading private records.')
