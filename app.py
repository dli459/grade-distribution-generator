import io
import hashlib
import re
import zipfile

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

@st.cache_data
def make_excel_template():
    columns = ['Course', 'Year', 'Term'] + GRADES
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        pd.DataFrame(columns=columns).to_excel(writer, sheet_name='Raw_Data', index=False)
        ws = writer.sheets['Raw_Data']
        ws.freeze_panes = 'D2'
        ws.auto_filter.ref = 'A1:O1'
        ws.column_dimensions['A'].width = 24
        ws.column_dimensions['B'].width = 13
        ws.column_dimensions['C'].width = 19
        for col in 'DEFGHIJKLMNO':
            ws.column_dimensions[col].width = 11
        from openpyxl.styles import PatternFill, Font, Alignment
        for cell in ws[1]:
            cell.fill = PatternFill('solid', fgColor='173A57')
            cell.font = Font(color='FFFFFF', bold=True)
            cell.alignment = Alignment(horizontal='center')
    return buffer.getvalue()

st.download_button(
    '⬇️ Download blank Excel template',
    data=make_excel_template(),
    file_name='Grade_Distribution_Input_Template.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    help='Fill one course / year / semester per row. Enter grade counts as non-negative whole numbers.',
)


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


@st.cache_data(show_spinner=False, max_entries=160)
def draw_chart_cached(selection, grade_values, kind):
    counts = np.array(grade_values, dtype=float)
    return _render_chart(selection, counts, kind)


def _render_chart(selection, counts, kind):
    total = counts.sum()
    values = counts if kind == 'Headcount' else counts / total * 100
    fig, ax = plt.subplots(figsize=(12.5, 5.5), dpi=160)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    bars = ax.bar(X, values, width=0.94, color=COLORS, edgecolor='none')
    ax.set_xticks(X, GRADES)
    ax.set_xlim(-0.65, 15.25)
    ax.set_ylabel('Count' if kind == 'Headcount' else 'Percentage (%)')
    ax.set_title(f"{selection} — {kind}", pad=17, fontsize=13, weight='bold')
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


def draw_chart(row, kind):
    return draw_chart_cached(str(row['Selection']), tuple(int(row[g]) for g in GRADES), kind)


def safe_name(value):
    return re.sub(r'[^A-Za-z0-9._-]+', '_', value).strip('_')[:100]

if mode == 'Selected course':
    row = df.loc[df['Selection'].eq(selected)].iloc[0]
    count_png = draw_chart(row, 'Headcount')
    pct_png = draw_chart(row, 'Percentage')
    left, right = st.columns(2)
    with left:
        st.image(count_png, caption='Headcount')
        st.download_button('Download Headcount PNG', count_png, f'{safe_name(selected)}_Headcount.png', 'image/png', on_click='ignore')
    with right:
        st.image(pct_png, caption='Percentage')
        st.download_button('Download Percentage PNG', pct_png, f'{safe_name(selected)}_Percentage.png', 'image/png', on_click='ignore')
    rows_to_export = [row]
else:
    rows_to_export = [row for _, row in df.iterrows()]
    st.write(f'Ready to export **{len(rows_to_export) * 2} PNG charts**.')

# Generate bulk ZIP only on request, and report actual chart-generation progress.
# Store completed bytes so Streamlit reruns do not regenerate every PNG.
file_signature = hashlib.sha256(file.getvalue()).hexdigest()
export_signature = (file_signature, mode, selected if mode == 'Selected course' else 'ALL')
if st.session_state.get('zip_export_signature') != export_signature:
    st.session_state.pop('zip_export_bytes', None)
    st.session_state['zip_export_signature'] = export_signature

chart_count = len(rows_to_export) * 2
if st.button(f'Generate ZIP ({chart_count} PNG charts)', type='primary', key='generate_zip'):
    progress_bar = st.progress(0, text=f'Preparing {chart_count} charts…')
    status = st.empty()
    zip_buffer = io.BytesIO()
    completed = 0
    try:
        with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            for row in rows_to_export:
                name = safe_name(row['Selection'])
                for kind in ('Headcount', 'Percentage'):
                    status.caption(f'Generating {completed + 1} of {chart_count}: {row["Selection"]} — {kind}')
                    archive.writestr(f'{name}_{kind}.png', draw_chart(row, kind))
                    completed += 1
                    progress_bar.progress(completed / chart_count,
                                          text=f'{completed} / {chart_count} charts generated ({completed / chart_count:.0%})')
        st.session_state['zip_export_bytes'] = zip_buffer.getvalue()
        progress_bar.empty()
        status.success(f'All {chart_count} charts are ready. Download the ZIP below.')
    except Exception as exc:
        st.session_state.pop('zip_export_bytes', None)
        status.error(f'Could not generate the ZIP: {exc}')

if 'zip_export_bytes' in st.session_state:
    st.download_button('📦 Download all charts as ZIP',
                       data=st.session_state['zip_export_bytes'],
                       file_name='Grade_Distribution_Charts.zip',
                       mime='application/zip', type='primary',
                       on_click='ignore')
else:
    st.caption('Click Generate ZIP to create the images. A progress bar will show each chart as it finishes.')


# Cross-course / cross-year A-B-C-D comparison (one bar per selected row).
st.divider()
st.header('Compare Courses / Years — 100% Stacked Percentage')
st.write('Filter by course, year and semester, then select the records you want in one comparison chart. Each bar sums to 100%.')

# Keep course/year/semester filters independent of the selected comparison rows.
# Reconcile stale widget state before constructing the multiselect, otherwise
# Streamlit may display selected pills that no longer belong to its options.
comparison_source = df.copy()
comparison_source['Course Name'] = comparison_source['Selection'].str.split(' | ', regex=False).str[0]
comparison_source['Year Label'] = comparison_source['Year'].astype(str).str.replace(r'\.0$', '', regex=True)
comparison_source['Term Label'] = comparison_source['Term'].astype(str).str.strip()

# A new upload must not inherit record selections from the previous workbook.
workbook_signature = file_signature
if st.session_state.get('_comparison_workbook_signature') != workbook_signature:
    for state_key in ('compare_courses', 'compare_years', 'compare_terms', 'compare_records', 'comparison_applied_filters'):
        st.session_state.pop(state_key, None)
    st.session_state['_comparison_workbook_signature'] = workbook_signature

course_options = sorted(comparison_source['Course Name'].unique().tolist())
year_options = sorted(comparison_source['Year Label'].unique().tolist())
term_options = sorted(comparison_source['Term Label'].unique().tolist())

with st.form('comparison_filter_form'):
    st.caption('Choose any combination of filters, then click Apply filters. Empty filters mean all records.')
    filter_columns = st.columns(3)
    with filter_columns[0]:
        course_filter = st.multiselect(
            'Course', course_options, default=[], key='compare_courses',
            help='Leave empty to include all courses.',
        )
    with filter_columns[1]:
        year_filter = st.multiselect(
            'Year', year_options, default=[], key='compare_years',
            help='Leave empty to include all years.',
        )
    with filter_columns[2]:
        term_filter = st.multiselect(
            'Semester / Term', term_options, default=[], key='compare_terms',
            help='Leave empty to include all semesters.',
        )
    apply_filters = st.form_submit_button('Apply filters', type='primary')

# The applied filters are separate from form widget state. Form edits do not
# trigger chart generation or rebuild the page until Apply filters is clicked.
filter_signature = (tuple(course_filter), tuple(year_filter), tuple(term_filter))
if apply_filters or 'comparison_applied_filters' not in st.session_state:
    st.session_state['comparison_applied_filters'] = filter_signature
course_filter, year_filter, term_filter = st.session_state['comparison_applied_filters']

comparison_df = comparison_source
if course_filter:
    comparison_df = comparison_df[comparison_df['Course Name'].isin(course_filter)]
if year_filter:
    comparison_df = comparison_df[comparison_df['Year Label'].isin(year_filter)]
if term_filter:
    comparison_df = comparison_df[comparison_df['Term Label'].isin(term_filter)]

options = comparison_df['Selection'].tolist()
# Clear invalid selections BEFORE the widget is created. This prevents the
# old 'No results' state when switching courses, years, or semesters.
if 'compare_records' in st.session_state:
    st.session_state['compare_records'] = [
        item for item in st.session_state['compare_records'] if item in options
    ]
else:
    st.session_state['compare_records'] = options[:10]

st.caption(f'{len(options)} matching record(s). Select the individual records to plot below.')
comparison_selections = st.multiselect(
    'Courses / semesters to include in the chart', options,
    key='compare_records',
    help='Each selected course/year/semester is a separate bar. Change the filters above to narrow the choices.',
)
if options:
    st.button('Select all matching records', key='select_all_comparison',
              on_click=lambda choices=options.copy(): st.session_state.update(compare_records=choices))
else:
    st.info('No records match these filters. Clear one or more filters to continue.')

STACK_LABELS = ['A Grades', 'B Grades', 'C Grades', 'D Grades']
STACK_COLORS = ['#22547c', '#315c9d', '#578638', '#c90000']

@st.cache_data(show_spinner=False, max_entries=32)
def draw_stacked_comparison_cached(selection_labels, grade_matrix, years, terms):
    rows = pd.DataFrame(grade_matrix, columns=GRADES)
    rows['Selection'] = list(selection_labels)
    rows['Year'] = list(years)
    rows['Term'] = list(terms)
    return draw_stacked_comparison(rows)

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
    comparison_png = draw_stacked_comparison_cached(
        tuple(comparison_rows['Selection'].astype(str)),
        tuple(tuple(int(v) for v in r) for r in comparison_rows[GRADES].to_numpy()),
        tuple(comparison_rows['Year'].astype(str)),
        tuple(comparison_rows['Term'].astype(str)),
    )
    st.image(comparison_png, caption=f'{len(comparison_selections)} selected course/semester records')
    st.download_button('Download comparison PNG', comparison_png,
                       'Grade_Distribution_Comparison_Stacked.png', 'image/png',
                       key='comparison_download', on_click='ignore')
else:
    st.info('Choose at least one course/semester record to create the comparison chart.')

st.caption('The uploaded workbook is processed in the app session; this prototype does not save it to a database. Hosting providers may still process uploaded data, so check your institution’s data policy before uploading private records.')
