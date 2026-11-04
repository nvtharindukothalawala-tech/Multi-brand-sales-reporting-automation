"""Local company reporting dashboard: read-only audit and verified copies.

Launch from repository root: python -m streamlit run streamlit_app.py
Requires existing src.real_business_dry_run and scripts.export_verified_channelwise_copies.
"""
import os
from pathlib import Path
import streamlit as st

from src.reporting_ui_service import (
    can_export_verified_copies,
    create_verified_copies_zip,
    input_fingerprint,
    review_rows,
    run_read_only_review,
)

from src.additional_transactions_report import create_additional_transactions_report_zip

st.set_page_config(page_title='Multi-Brand Reporting', page_icon='📊', layout='wide')
st.title('Multi-Brand ChannelWise Reporting')
st.caption('Private local dashboard | Real-data read-only validation + verified copies')
st.info('This version NEVER inserts transactions into real Excel workbooks. New or changed transactions require manual review and an approved update engine.')

with st.sidebar:
    st.header('Monthly report settings')
    month = st.text_input('Reporting month (YYYY-MM)', value='2026-10')
    templates_dir = st.text_input(
        'Original ChannelWise template folder',
        value=os.environ.get('CHANNELWISE_TEMPLATES_DIR', str(Path.home() / 'Desktop' / 'Ass')),
    )
    st.caption('Keep five original workbooks in a private local folder, outside GitHub.')

left, right = st.columns(2)
with left:
    sales = st.file_uploader('Upload monthly sales export (.xlsx)', type=['xlsx'], key='sales_input')
with right:
    returns = st.file_uploader('Upload monthly returns export (.xlsx)', type=['xlsx'], key='returns_input')

if 'review_state' not in st.session_state:
    st.session_state.review_state = None
if 'download_state' not in st.session_state:
    st.session_state.download_state = None
if 'additional_download_state' not in st.session_state:
    st.session_state.additional_download_state = None

ready = sales is not None and returns is not None
sales_bytes = sales.getvalue() if sales else b''
returns_bytes = returns.getvalue() if returns else b''
try:
    fingerprint = input_fingerprint(sales_bytes, returns_bytes, month, templates_dir) if ready else None
except (ValueError, OSError):
    fingerprint = None

if st.button('1. Validate monthly data (read-only)', type='primary', disabled=not ready):
    st.session_state.download_state = None
    st.session_state.additional_download_state = None
    st.session_state.review_state = None
    try:
        with st.spinner('Comparing original exports against all five company workbooks...'):
            report = run_read_only_review(sales_bytes, returns_bytes, month, templates_dir)
        st.session_state.review_state = report
    except Exception as exc:
        st.error(f'Validation failed: {exc}')

report = st.session_state.review_state
if report is not None and report.get('ui_input_fingerprint') == fingerprint:
    st.subheader('Validation results')
    st.metric('Audit status', report['status'])
    a, b, c = st.columns(3)
    a.metric('Mapped sales', report['mapped_sales_records'])
    b.metric('Mapped returns', report['mapped_return_records'])
    c.metric('Unmapped records', report['unmapped_sales']['records'] + report['unmapped_returns']['records'])
    st.dataframe(review_rows(report), width='stretch', hide_index=True)
    if report.get('warnings'):
        with st.expander('Warnings and scope review', expanded=True):
            for warning in report['warnings']:
                st.warning(warning)
    st.download_button(
        'Download private audit JSON',
        data=__import__('json').dumps(report, indent=2),
        file_name=f'PRIVATE_channelwise_audit_{month}.json',
        mime='application/json',
    )
    if st.button('Prepare additional transactions report (ZIP)'):
        try:
            with st.spinner('Preparing confidential workbook-only transaction details...'):
                additional_zip = create_additional_transactions_report_zip(
                    sales_bytes, returns_bytes, month, templates_dir, report,
                )
            st.session_state.additional_download_state = (fingerprint, additional_zip)
            st.success('Additional transaction report is ready. No workbook changes were made.')
        except Exception as exc:
            st.session_state.additional_download_state = None
            st.error(f'Additional report blocked: {exc}')
    additional_saved = st.session_state.additional_download_state
    if additional_saved is not None and additional_saved[0] == fingerprint:
        st.download_button(
            'Download all additional transactions (ZIP)',
            data=additional_saved[1],
            file_name=f'PRIVATE_Additional_Transactions_{month}.zip',
            mime='application/zip',
        )
    approved = st.checkbox('I understand these are identical report copies, not updated monthly reports.')
    allowed = can_export_verified_copies(report)
    if not allowed:
        st.error('Report export is blocked: the source contains missing or conflicting mapped transactions. No workbook updates are permitted.')
    if st.button('2. Prepare verified copies (ZIP)', disabled=not (allowed and approved)):
        try:
            with st.spinner('Revalidating and copying five workbooks without changing originals...'):
                zip_bytes = create_verified_copies_zip(sales_bytes, returns_bytes, month, templates_dir, report)
            st.session_state.download_state = (fingerprint, zip_bytes)
            st.success('Five verified workbook copies are ready. No transactions were inserted.')
        except Exception as exc:
            st.session_state.download_state = None
            st.error(f'Export blocked: {exc}')
    saved = st.session_state.download_state
    if saved is not None and saved[0] == fingerprint and approved:
        st.download_button(
            '3. Download five ChannelWise reports (ZIP)',
            data=saved[1],
            file_name=f'PRIVATE_ChannelWise_verified_copies_{month}.zip',
            mime='application/zip',
            type='primary',
        )
elif ready:
    st.caption('Select “Validate monthly data” to check sales and returns against existing templates.')

st.divider()
st.caption('Safety: originals are read-only. No production updates, Excel formula recalculation, or pivot refresh. Do not deploy this app publicly with private company data.')
