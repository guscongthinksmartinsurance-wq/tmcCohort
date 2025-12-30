import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
from io import BytesIO

# --- 1. GIAO DIỆN ---
st.set_page_config(page_title="TMC Strategic Portal", layout="wide")
st.markdown("""
    <style>
    .main { background-color: #0E1117; color: #FFFFFF; }
    [data-testid="stMetricValue"] { color: #00D4FF !important; font-weight: 900 !important; font-size: 2.5rem !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. HÀM ĐỌC FILE ---
def smart_load(file):
    try:
        if file.name.endswith(('.xlsx', '.xls')):
            raw_df = pd.read_excel(file, header=None)
        else:
            file.seek(0)
            raw_df = pd.read_csv(file, sep=None, engine='python', header=None, encoding='utf-8', errors='ignore')
        header_row = 0
        for i, row in raw_df.head(20).iterrows():
            if 'TARGET PREMIUM' in " ".join(str(val).upper() for val in row):
                header_row = i; break
        file.seek(0)
        return pd.read_excel(file, skiprows=header_row) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, sep=None, engine='python', skiprows=header_row, encoding='utf-8', errors='ignore')
    except: return None

# --- 3. ENGINE PHÂN TÍCH ---
def process_data(file):
    df = smart_load(file)
    if df is None: return

    curr_y = datetime.now().year
    cols = df.columns
    c_list = [" ".join(str(c).upper().split()) for c in cols]
    
    def get_c(keys):
        for i, c in enumerate(c_list):
            if all(k in c for k in keys): return cols[i]
        return None

    m_c, e_c, v_c, w_c, id_c, src_c = get_c(['TARGET', 'PREMIUM']), get_c(['THÁNG', 'NHẬN', 'FILE']), get_c(['THÁNG', 'NHẬN', 'LEAD']), get_c(['NĂM', 'NHẬN', 'LEAD']), get_c(['LEAD', 'ID']), get_c(['SOURCE'])

    # Làm sạch Doanh thu
    df['REV'] = df[m_c].apply(lambda v: float(re.sub(r'[^0-9.]', '', str(v))) if pd.notna(v) and re.sub(r'[^0-9.]', '', str(v)) != '' else 0.0)

    # Logic Phân loại Nhóm (Linh hoạt theo năm)
    def assign_cohort(row):
        try:
            # Nhận diện Cold Call trước
            if src_c and str(row[src_c]).strip().upper() in ['CC', 'COLDCALL']:
                return "📞 Kênh Cold Call"
            
            # Xử lý năm và tháng
            y = int(float(str(row[w_c]).strip()))
            m = int(float(str(row[v_c]).strip()))
            return f"Lead T{m:02d}/{y}"
        except:
            return "📦 Dữ liệu chưa phân loại"

    df['NHÓM_LEAD'] = df.apply(assign_cohort, axis=1)
    df['TH_CHOT_NUM'] = df[e_c].apply(lambda v: int(float(v)) if pd.notna(v) and 1 <= int(float(v)) <= 12 else None)

    # Ma trận Doanh số & Số lượng
    matrix_rev = df.pivot_table(index='NHÓM_LEAD', columns='TH_CHOT_NUM', values='REV', aggfunc='sum').fillna(0)
    matrix_count = df.pivot_table(index='NHÓM_LEAD', columns='TH_CHOT_NUM', values=id_c, aggfunc='nunique').fillna(0)

    # Sắp xếp Matrix theo thứ tự thời gian giảm dần (Mới nhất lên đầu)
    def sort_mtx(mtx):
        mtx = mtx.reindex(columns=range(1, 13)).fillna(0)
        mtx.columns = [f"Tháng {int(c)}" for c in mtx.columns]
        
        # Tách các nhóm để sort
        idx_cc = [i for i in mtx.index if "Cold Call" in i]
        idx_miss = [i for i in mtx.index if "chưa phân loại" in i]
        idx_leads = [i for i in mtx.index if "Lead T" in i]
        
        # Sort Lead theo năm giảm dần, tháng giảm dần
        idx_leads.sort(key=lambda x: (int(x.split('/')[-1]), int(x.split('T')[-1].split('/')[0])), reverse=True)
        
        return mtx.reindex(idx_leads + idx_cc + idx_miss)

    matrix_rev = sort_mtx(matrix_rev)
    matrix_count = sort_mtx(matrix_count)

    # --- HIỂN THỊ ---
    st.title(f"🚀 Strategic Growth Analysis - {curr_y}")
    
    col1, col2, col3 = st.columns(3)
    total_rev = df['REV'].sum()
    marketing_rev = df[df['NHÓM_LEAD'].str.contains('Lead')]['REV'].sum()
    
    col1.metric("💰 TỔNG DOANH THU", f"${total_rev:,.0f}")
    col2.metric("🎯 DOANH THU MARKETING", f"${marketing_rev:,.0f}")
    col3.metric("📋 TỔNG HỒ SƠ", f"{df[id_c].nunique():,}")

    t1, t2 = st.tabs(["💵 Ma trận Doanh số ($)", "🔢 Ma trận Số lượng (Hồ sơ)"])
    with t1: st.dataframe(matrix_rev.style.format("${:,.0f}"), use_container_width=True)
    with t2: st.dataframe(matrix_count.style.format("{:,.0f}"), use_container_width=True)

    # Xuất Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        matrix_rev.to_excel(writer, sheet_name='Revenue_Cohort')
        matrix_count.to_excel(writer, sheet_name='Count_Cohort')
        df.to_excel(writer, index=False, sheet_name='Full_Data')
    st.sidebar.download_button("📥 Tải Báo Cáo Strategic", output.getvalue(), f"TMC_Report_{curr_y}.xlsx")

st.title("🛡️ Strategic Portal")
f = st.file_uploader("Nạp dữ liệu Masterlife", type=['csv', 'xlsx'])
if f: process_data(f)
