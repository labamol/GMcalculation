"""Run with: streamlit run app.py"""
import json
from datetime import date, datetime
from decimal import Decimal

import pandas as pd
import streamlit as st

from gmcalc.core import ValidationError, baseline, headcount, margin, margin_from_workbooks

st.set_page_config(page_title='Gross Margin Calculation', layout='wide')
st.title('Gross Margin Calculation')
st.caption('Quarterly revenue in INR Crores · monthly employee cost in INR · quarterly cost projection ×3')
st.write('Upload original .xlsx workbooks. Processing runs locally in this application. Inputs are never overwritten. Exact, case-sensitive Segment and Account/Customer matches are used after trimming whitespace.')


def json_default(value):
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    raise TypeError(type(value).__name__)


def display(result, filename, widget_key):
    st.subheader(filename)
    if result.status == 'PASS':
        st.success('Data checks PASS')
    else:
        st.warning('Validation FAIL / review required. Inspect exceptions before using financial results.')
    limitation = result.metadata.get('formula_validation_limitation')
    if limitation:
        st.warning(limitation)
    st.download_button('Download Excel', result.workbook, filename,
                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key=widget_key)
    report = json.dumps(result.report(), indent=2, default=json_default)
    st.download_button('Download validation report', report, filename.replace('.xlsx', '_validation.json'),
                       'application/json', key=widget_key + '_report')
    with st.expander('Validation and source metadata', expanded=result.status != 'PASS'):
        st.json(json.loads(report))
    for name, table in result.tables.items():
        with st.expander(name, expanded=name != 'Key Data'):
            df = pd.DataFrame(table['rows'], columns=table['headers'])
            # Display numbers as numbers; Excel exports retain underlying precision.
            for col in df:
                df[col] = df[col].map(lambda x: float(x) if isinstance(x, Decimal) else x)
            st.dataframe(df, hide_index=True, width='stretch')


mode = st.radio('Workflow', ['All steps', 'Step 1: Revenue', 'Step 2: Headcount', 'Step 3: Margin'], horizontal=True)
revenue_upload = None
hc_upload = None
if mode != 'Step 2: Headcount':
    revenue_upload = st.file_uploader('Quarterly baseline workbook', type=['xlsx'], key='revenue')
    quarter = st.selectbox('Reporting quarter (must match sheet and column meanings)', ['Q3', 'Q1', 'Q2', 'Q4'])
    revenue_header = st.number_input('Baseline last header row (all following rows are detail)', min_value=1, value=1)
    excluded = st.text_input('Verified subtotal/heading row numbers to exclude', placeholder='e.g. 12, 25, 40')
    blank_zero = st.checkbox('I confirm blank revenue cells represent missing contribution and may be zero')
if mode != 'Step 1: Revenue':
    hc_upload = st.file_uploader('Daily headcount workbook', type=['xlsx'], key='headcount')
    hc_header = st.number_input('Data sheet header row', min_value=1, value=1)
verified = st.checkbox('If source formulas exist, I confirm these files were recalculated and saved in Excel/LibreOffice immediately before upload')
if mode == 'Step 3: Margin':
    st.info('Step 3 reads refreshed Segment View, Account View and Costs By Segment sheets. Header-row, exclusion and blank-revenue options above apply only to Steps 1–2. Resolve upstream validation exceptions before using refreshed files; these files do not embed their JSON validation reports.')
st.caption('Formula caches cannot be dated reliably by this application. Missing caches are always rejected. Only app-marked generated sheets can be refreshed. Workbooks with unsupported preservation features are rejected.')

if st.button('Process selected inputs', type='primary'):
    st.session_state.pop('results', None)
    try:
        results = []
        rev = hc = None
        if mode == 'Step 3: Margin':
            if revenue_upload is None or hc_upload is None:
                raise ValidationError('Upload both refreshed workbooks containing Segment View, Account View and Costs By Segment')
            result = margin_from_workbooks(revenue_upload.getvalue(), hc_upload.getvalue(), quarter,
                                           verified, revenue_upload.name, hc_upload.name)
            results.append((result, 'Margin Calculation.xlsx', 'margin_out'))
        if mode not in ('Step 2: Headcount', 'Step 3: Margin'):
            if revenue_upload is None:
                raise ValidationError('Upload a quarterly baseline workbook')
            try:
                excluded_rows = [int(x.strip()) for x in excluded.split(',') if x.strip()]
            except ValueError as exc:
                raise ValidationError('Excluded rows must be comma-separated integer row numbers') from exc
            rev = baseline(revenue_upload.getvalue(), quarter, int(revenue_header), excluded_rows,
                           blank_zero, verified, revenue_upload.name)
            results.append((rev, 'Quarterly Baseline Updated.xlsx', 'rev_out'))
        if mode not in ('Step 1: Revenue', 'Step 3: Margin'):
            if hc_upload is None:
                raise ValidationError('Upload a daily headcount workbook')
            hc = headcount(hc_upload.getvalue(), int(hc_header), verified, hc_upload.name)
            results.append((hc, 'Headcount Updated.xlsx', 'hc_out'))
        if rev is not None and hc is not None:
            results.append((margin(rev, hc), 'Margin Calculation.xlsx', 'margin_out'))
        st.session_state['results'] = results
    except (ValidationError, ValueError, OSError) as exc:
        st.error(str(exc))

if 'results' in st.session_state:
    st.info('Results below reflect the last successful processing run. After changing inputs or settings, click Process again. Excel filter-aware cost subtotals require desktop Excel/LibreOffice recalculation and filter validation.')
    for result, filename, widget_key in st.session_state['results']:
        display(result, filename, widget_key)
