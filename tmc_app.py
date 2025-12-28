import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
from io import BytesIO

# --- 1. CSS GIAO DIỆN CAO CẤP & TINH TẾ ---
st.set_page_config(page_title="TMC Strategic Portal", layout="wide")
st.markdown("""
    <style>
    .main { background-color: #0E1117; color: #FFFFFF; }
    [data-testid="stMetricValue"] { color: #00D4FF !important; font-weight: 900 !important; font-size: 2.5rem !important; }
    [data-testid="stMetricLabel"] p { color: #8B949E !important; font-size: 0.9rem !important; letter-spacing: 1px; }
    [data-testid="stChart"] { height: 350px !important; }
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
                header_row = i
                break
        file.seek(0)
        return pd.read_excel(file, skiprows=header_row) if file.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file, sep=None, engine='python', skiprows=header_row, encoding='utf-8', errors='ignore')
    except: return None

# --- 3. ENGINE PHÂN TÍCH ---
def process_data(file):
    df = smart_load(file)
    if df is None: return

    current_year = datetime.now().year
    cols = df.columns
    c_list = [" ".join(str(c).upper().split()) for c in cols]
    def get_c(keys):
        for i, c in enumerate(c_list):
            if all(k in c for k in keys): return cols[i]
        return None

    m_c, e_c, v_c, w_c, id_c, src_c = get_c(['TARGET', 'PREMIUM']), get_c(['THÁNG', 'NHẬN', 'FILE']), get_c(['THÁNG', 'NHẬN', 'LEAD']), get_c(['NĂM', 'NHẬN', 'LEAD']), get_c(['LEAD', 'ID']), get_c(['SOURCE'])

    # Làm sạch & Lọc Funnel
    if src_c:
        df = df.dropna(subset=[src_c])
        df['SRC_UP'] = df[src_c].astype(str).str.upper().str.replace(" ", "").str.replace(".", "")
        df = df[~df['SRC_UP'].isin(['CC', 'COLDCALL'])]

    df['REV'] = df[m_c].apply(lambda v: float(re.sub(r'[^0-9.]', '', str(v))) if pd.notna(v) and re.sub(r'[^0-9.]', '', str(v)) != '' else 0.0)
    
    def assign_cohort(row):
        try:
            y, m = int(float(row[w_c])), int(float(row[v_c]))
            return f"Lead T{m:02d}/{y}" if y == current_year else f"Trước năm {current_year}"
        except: return "❌ Thiếu thông tin Lead"

    df['NHÓM_LEAD'] = df.apply(assign_cohort, axis=1)
    df['TH_CHOT_NUM'] = df[e_c].apply(lambda v: int(float(v)) if pd.notna(v) and 1 <= int(float(v)) <= 12 else None)

    # Dữ liệu biểu đồ
    full_year_data = df.groupby('TH_CHOT_NUM')['REV'].sum().reindex(range(1, 13)).fillna(0)
    chart_df = pd.DataFrame({'Tháng': [f"Tháng {i:02d}" for i in range(1, 13)], 'Doanh Số': full_year_data.values}).set_index('Tháng')

    # Ma trận Doanh số & Số lượng
    matrix_rev = df.pivot_table(index='NHÓM_LEAD', columns='TH_CHOT_NUM', values='REV', aggfunc='sum').fillna(0)
    matrix_count = df.pivot_table(index='NHÓM_LEAD', columns='TH_CHOT_NUM', values=id_c, aggfunc='nunique').fillna(0)

    def sort_mtx(mtx):
        mtx = mtx.reindex(columns=range(1, 13)).fillna(0)
        mtx.columns = [f"Tháng {int(c)}" for c in mtx.columns]
        idx_current = sorted([i for i in mtx.index if f"/{current_year}" in i])
        final_idx = ([f"Trước năm {current_year}"] if f"Trước năm {current_year}" in mtx.index else []) + idx_current + ([i for i in mtx.index if "❌" in i])
        return mtx.reindex(final_idx)

    matrix_rev = sort_mtx(matrix_rev)
    matrix_count = sort_mtx(matrix_count)

    # --- HIỂN THỊ ---
    st.title(f"🚀 Strategic Growth Analysis - {current_year}")
    st.area_chart(chart_df, color="#00D4FF")

    col1, col2 = st.columns(2)
    col1.metric("💰 TỔNG DOANH THU", f"${df['REV'].sum():,.2f}")
    col2.metric("📋 HỒ SƠ HOÀN TẤT", f"{df[id_c].nunique():,}")

    t1, t2 = st.tabs(["💵 Ma trận Doanh số ($)", "🔢 Ma trận Số lượng (Hồ sơ)"])
    with t1: st.dataframe(matrix_rev.style.format("${:,.0f}"), use_container_width=True)
    with t2: st.dataframe(matrix_count.style.format("{:,.0f}"), use_container_width=True)

    # --- XUẤT EXCEL ĐA SHEET ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet 1: Doanh số
        matrix_rev.to_excel(writer, sheet_name='Summary_Cohort')
        ws1 = writer.sheets['Summary_Cohort']
        ws1.conditional_format(1, 1, len(matrix_rev), 12, {'type': '3_color_scale', 'min_color': "#F7FBFF", 'mid_color': "#6BAED6", 'max_color': "#08306B"})
        
        # Sheet 2: Số lượng (Bị thiếu lúc trước)
        matrix_count.to_excel(writer, sheet_name='Count_Cohort')
        ws2 = writer.sheets['Count_Cohort']
        ws2.conditional_format(1, 1, len(matrix_count), 12, {'type': '3_color_scale', 'min_color': "#FFF5EB", 'mid_color': "#FDAE6B", 'max_color': "#7F2704"})
        
        # Sheet 3: Data nguồn
        df.to_excel(writer, index=False, sheet_name='Full_Clean_Data')

    st.sidebar.markdown("---")
    st.sidebar.download_button("📥 Tải Báo Cáo Strategic (.xlsx)", output.getvalue(), f"TMC_Strategic_Report_{current_year}.xlsx")

st.title("🛡️ Strategic Portal")
f = st.file_uploader("Nạp dữ liệu Masterlife (CSV/Excel)", type=['csv', 'xlsx'])
if f: process_data(f)