import streamlit as st
import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

st.set_page_config(page_title="Daily Payment Sheet", page_icon="💳", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .header-block {
        background: linear-gradient(135deg, #1a2f5e 0%, #2d4a8a 100%);
        border-radius: 16px; padding: 32px 36px; margin-bottom: 28px;
    }
    .header-block h1 { font-size: 1.8rem; font-weight: 700; margin: 0 0 6px 0; color: white; }
    .header-block p  { font-size: 0.95rem; opacity: 0.8; margin: 0; color: white; }
    .instructions {
        background: #ffffff; border-radius: 12px; padding: 20px 24px;
        border-left: 4px solid #2d4a8a; margin-bottom: 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .instructions h4 {
        margin: 0 0 10px 0; color: #1a2f5e; font-size: 0.95rem;
        font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em;
    }
    .instructions ol { margin: 0; padding-left: 18px; color: #444; font-size: 0.92rem; line-height: 1.8; }
    .result-box {
        background: #f0f7f0; border: 1px solid #b2dfb2; border-radius: 12px;
        padding: 20px 24px; margin-top: 20px; text-align: center;
    }
    .result-box h3 { color: #2e7d32; margin: 0 0 8px 0; font-size: 1.1rem; }
    .result-box p  { color: #555; margin: 0; font-size: 0.9rem; }
    .stat-row { display: flex; gap: 12px; margin: 16px 0; justify-content: center; }
    .stat-card {
        background: white; border-radius: 10px; padding: 14px 20px;
        text-align: center; border: 1px solid #e0e0e0; min-width: 120px;
    }
    .stat-card .number { font-size: 1.6rem; font-weight: 700; color: #1a2f5e; }
    .stat-card .label  { font-size: 0.78rem; color: #888; text-transform: uppercase; letter-spacing: 0.04em; margin-top: 2px; }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #1a2f5e, #2d4a8a) !important;
        color: white !important; border: none !important; border-radius: 10px !important;
        padding: 12px 32px !important; font-size: 1rem !important;
        font-weight: 600 !important; width: 100% !important; margin-top: 12px !important;
    }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-block">
    <h1>💳 Daily Payment Sheet</h1>
    <p>Lissa D Mills Inc — Front Desk Payment Processing</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="instructions">
    <h4>How to use</h4>
    <ol>
        <li>Run your daily payment report in Practice Pro and save the <strong>.xls</strong> file</li>
        <li>Upload the file below</li>
        <li>Click <strong>Generate Payment Sheet</strong></li>
        <li>Download the finished Excel file — ready to print or use at the front desk</li>
    </ol>
</div>
""", unsafe_allow_html=True)


def process_payment_sheet(uploaded_file):
    df = pd.read_excel(uploaded_file, engine='xlrd', header=None)

    header_row = None
    for i, row in df.iterrows():
        if 'Appt. Time' in str(list(row.values)):
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not find 'Appt. Time' header. Please check you uploaded the correct Practice Pro report.")

    cols = df.iloc[header_row].values
    data = df.iloc[header_row + 1:].copy()
    data.columns = range(len(data.columns))
    col_map = {str(val).strip(): idx for idx, val in enumerate(cols)}

    appt_col = col_map.get('Appt. Time', 1)
    name_col = col_map.get('Patient Name', 2)
    fee_col  = col_map.get('Visit Fee', 6)

    data = data[data[appt_col].apply(lambda x: isinstance(x, datetime.datetime))].copy()
    data['visit_fee']    = pd.to_numeric(data[fee_col], errors='coerce').fillna(0)
    data['patient_name'] = data[name_col].astype(str).str.strip()
    data['appt_time']    = data[appt_col]
    data = data[data['visit_fee'] > 0]

    grouped = (
        data.groupby('patient_name', sort=False)
        .agg(visit_fee=('visit_fee', 'sum'), first_appt=('appt_time', 'min'))
        .reset_index()
        .sort_values('first_appt')
        .reset_index(drop=True)
    )

    report_date  = data['appt_time'].min().date()
    total_amount = grouped['visit_fee'].sum()

    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Payment Sheet"

    ws.merge_cells('A1:G1')
    ws['A1'].value     = "Daily Payment Sheet Report"
    ws['A1'].font      = Font(name='Arial', bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A2:G2')
    ws['A2'].value     = report_date.strftime('%A, %B %d, %Y')
    ws['A2'].font      = Font(name='Arial', size=11)
    ws['A2'].alignment = Alignment(horizontal='center')

    ws.append([])

    thin_black     = Side(style='thin',   color='000000')
    med_black      = Side(style='medium', color='000000')
    full_border    = Border(top=thin_black, bottom=thin_black, left=thin_black, right=thin_black)
    header_border  = Border(top=thin_black, bottom=med_black,  left=thin_black, right=thin_black)
    summary_border = Border(top=med_black,  bottom=med_black,  left=thin_black, right=thin_black)

    headers     = ['Patient Name', 'Visit Fee', 'Cash', 'Credit', 'Check', 'Notes', 'Posted in PPro']
    header_fill = PatternFill('solid', start_color='1a2f5e', end_color='1a2f5e')
    header_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)

    ws.append(headers)
    for col_idx in range(1, 8):
        cell           = ws.cell(row=4, column=col_idx)
        cell.fill      = header_fill
        cell.font      = header_font
        cell.border    = header_border
        cell.alignment = Alignment(horizontal='left' if col_idx == 1 else 'center')

    for i, row in grouped.iterrows():
        row_num = 5 + i

        name_cell      = ws.cell(row=row_num, column=1, value=row['patient_name'])
        name_cell.font = Font(name='Arial', size=10)

        fee_cell               = ws.cell(row=row_num, column=2, value=row['visit_fee'])
        fee_cell.number_format = '$#,##0.00'
        fee_cell.font          = Font(name='Arial', size=10)
        fee_cell.alignment     = Alignment(horizontal='right')

        for col_idx in range(3, 7):
            ws.cell(row=row_num, column=col_idx, value='')

        posted_cell           = ws.cell(row=row_num, column=7, value='')
        posted_cell.alignment = Alignment(horizontal='center', vertical='center')
        posted_cell.fill      = PatternFill('solid', start_color='E8F5E9', end_color='E8F5E9')

        if i % 2 == 1:
            for col_idx in range(1, 7):
                ws.cell(row=row_num, column=col_idx).fill = PatternFill('solid', start_color='F5F5F5', end_color='F5F5F5')
            ws.cell(row=row_num, column=7).fill = PatternFill('solid', start_color='DCEDC8', end_color='DCEDC8')

        for col_idx in range(1, 8):
            ws.cell(row=row_num, column=col_idx).border = full_border

    summary_row  = 5 + len(grouped)
    summary_fill = PatternFill('solid', start_color='1a2f5e', end_color='1a2f5e')
    summary_font = Font(name='Arial', bold=True, color='FFFFFF', size=10)

    ws.cell(row=summary_row, column=1, value="{} Patients".format(len(grouped)))
    fee_sum = ws.cell(row=summary_row, column=2, value=total_amount)
    fee_sum.number_format = '$#,##0.00'
    for col_idx in range(1, 8):
        c        = ws.cell(row=summary_row, column=col_idx)
        c.fill   = summary_fill
        c.font   = summary_font
        c.border = summary_border

    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 30
    ws.column_dimensions['G'].width = 14

    ws.freeze_panes           = 'A5'
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.print_title_rows       = '4:4'

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, len(grouped), total_amount, report_date


uploaded_file = st.file_uploader(
    "Upload your Practice Pro daily payment report (.xls)",
    type=['xls'],
    help="Export the Daily Payment Slip report from Practice Pro and upload it here"
)

if uploaded_file:
    if st.button("⚡ Generate Payment Sheet", use_container_width=True):
        with st.spinner("Processing your report..."):
            try:
                buffer, patient_count, total_amount, report_date = process_payment_sheet(uploaded_file)
                filename = "Daily_Payment_Sheet_{}.xlsx".format(report_date.strftime('%Y-%m-%d'))

                st.markdown("""
                <div class="result-box">
                    <h3>✅ Payment Sheet Ready!</h3>
                    <p>Your daily sheet has been generated and is ready to download.</p>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("""
                <div class="stat-row">
                    <div class="stat-card"><div class="number">{}</div><div class="label">Patients</div></div>
                    <div class="stat-card"><div class="number">${:,.2f}</div><div class="label">Total to Collect</div></div>
                    <div class="stat-card"><div class="number">{}</div><div class="label">Report Date</div></div>
                </div>
                """.format(patient_count, total_amount, report_date.strftime('%b %d')), unsafe_allow_html=True)

                st.download_button(
                    label="⬇️  Download Daily Payment Sheet",
                    data=buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error("❌ Error processing file: {}".format(str(e)))
                st.info("Please make sure you uploaded the correct Practice Pro Daily Payment Slip report (.xls format).")
else:
    st.markdown("""
    <div style="text-align:center; padding: 32px; color: #aaa; font-size: 0.9rem;">
        ⬆️ Upload your .xls file above to get started
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")
st.markdown("<div style='text-align:center; color:#bbb; font-size:0.8rem;'>Lissa D Mills Inc · Daily Payment Processing</div>", unsafe_allow_html=True)
