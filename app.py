import streamlit as st
import pandas as pd
import altair as alt
import os

# --- 1. ตั้งค่าหน้าเว็บ ---
st.set_page_config(
    page_title="BRIA NIPT Executive Report 2025",
    layout="wide",
    initial_sidebar_state="collapsed" # ซ่อน Sidebar เพราะไม่ต้องอัปโหลดแล้ว
)

# --- CSS ตกแต่ง ---
st.markdown("""
<style>
    .metric-card { background-color: #f9f9f9; padding: 15px; border-radius: 10px; border-left: 5px solid #2E7D32; box-shadow: 2px 2px 5px rgba(0,0,0,0.1); }
    .insight-box { background-color: #e8f5e9; padding: 15px; border-radius: 8px; margin-bottom: 10px; border: 1px solid #c8e6c9; }
    h1, h2, h3 { color: #1565C0; }
</style>
""", unsafe_allow_html=True)

# --- 2. ส่วนโหลดข้อมูล (กำหนดชื่อไฟล์ตายตัวตรงนี้) ---
# ⚠️ สำคัญ: คุณต้องวางไฟล์ Excel ไว้ที่เดียวกับไฟล์ Code นี้
# และตั้งชื่อไฟล์ให้ตรงกัน (ในที่นี้ผมสมมติว่าชื่อ 'data.xlsx')
DATA_FILENAME = "สรุปต้นทุน BRIA NIPT รายเดือน ปี 2025.xlsx" 

@st.cache_data
def load_data():
    # ตรวจสอบว่ามีไฟล์อยู่จริงไหม
    if not os.path.exists(DATA_FILENAME):
        return None

    xls = pd.ExcelFile(DATA_FILENAME)
    all_data = []
    
    for sheet_name in xls.sheet_names:
        # อ่านข้อมูล
        df_sheet = pd.read_excel(DATA_FILENAME, sheet_name=sheet_name)
        
        # เช็คคอลัมน์ (กันเหนียว เผื่อไปอ่านเจอ sheet สรุป)
        required_cols = ['Sales', 'NIPT Package', 'Gain', 'TAT']
        if not all(col in df_sheet.columns for col in required_cols):
            continue

        # จัดการชื่อเดือน
        found_month = sheet_name
        for m in ["May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March", "April"]:
            if m.lower() in sheet_name.lower():
                found_month = m
                break
        df_sheet['Month'] = found_month
        all_data.append(df_sheet)

    if not all_data:
        return pd.DataFrame()

    df_all = pd.concat(all_data, ignore_index=True)
    
    # Cleaning
    for col in ['Cost', 'Price', 'Gain', 'TAT']:
        df_all[col] = pd.to_numeric(df_all[col], errors='coerce')
    
    df_all['Sales'] = df_all['Sales'].fillna('Unknown')
    df_all = df_all.dropna(subset=['NIPT Package'])
    
    # เรียงเดือน
    month_order = ["May", "June", "July", "August", "September", "October", "November", "December"]
    existing_months = [m for m in month_order if m in df_all['Month'].unique()]
    df_all['Month'] = pd.Categorical(df_all['Month'], categories=existing_months, ordered=True)
    
    return df_all

# --- 3. เริ่มทำงาน ---
df = load_data()

if df is None:
    st.error(f"❌ ไม่พบไฟล์ข้อมูล: '{DATA_FILENAME}'")
    st.warning("กรุณานำไฟล์ Excel มาวางไว้ในโฟลเดอร์เดียวกับไฟล์ Code นี้ แล้วตั้งชื่อให้ตรงกันครับ")
    st.stop()
elif df.empty:
    st.error("❌ ไฟล์ Excel ไม่มีข้อมูลที่ถูกต้อง หรือรูปแบบคอลัมน์ไม่ตรง")
    st.stop()

# --- 4. คำนวณตัวเลข KPI ---
total_samples = len(df)
total_gain = df['Gain'].sum()
avg_tat = df['TAT'].mean()
monthly_gain = df.groupby('Month')['Gain'].sum()
best_month = monthly_gain.idxmax()
best_month_gain = monthly_gain.max()

# --- 5. แสดงผล Dashboard ---
st.title(f"🚀 BRIA NIPT Executive Dashboard")
st.markdown(f"**ข้อมูล Update ล่าสุด:** {df['Month'].max()} 2025")

# KPI Cards
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Cases", f"{total_samples:,.0f}", "สะสม")
col2.metric("Total Profit", f"{total_gain/1000000:,.2f} MB", f"฿{total_gain:,.0f}")
col3.metric("Avg TAT", f"{avg_tat:.1f} Days", "Target < 5")
col4.metric("Best Month", f"{best_month}", f"฿{best_month_gain:,.0f}")

st.markdown("---")

# Tabs
tab1, tab2 = st.tabs(["📊 Interactive Dashboard", "📝 Executive Summary"])

with tab1:
    # Selector
    selection = alt.selection_point(fields=['Month'])
    
    # Chart 1: Monthly Overview
    chart_main = alt.Chart(df).mark_bar().encode(
        x=alt.X('Month', title='Month'),
        y=alt.Y('count()', title='Number of Cases'),
        color=alt.condition(selection, alt.value('#1976D2'), alt.value('lightgray')),
        tooltip=['Month', 'count()', 'sum(Gain)']
    ).add_params(selection).properties(
        title='Monthly Volume (Click bar to filter)', height=300
    )
    
    # Chart 2: Product Mix
    chart_donut = alt.Chart(df).mark_arc(innerRadius=60).encode(
        theta=alt.Theta("count()", stack=True),
        color=alt.Color("NIPT Package", scale=alt.Scale(scheme='set2')),
        tooltip=["NIPT Package", "count()", alt.Tooltip("count()", format=".1%")]
    ).transform_filter(selection).properties(title='Product Mix', height=300)
    
    # Chart 3: Top Sales
    chart_sales = alt.Chart(df).mark_bar().encode(
        y=alt.Y('Sales', sort='-x'),
        x=alt.X('count()'),
        color=alt.value('#FF8F00'),
        tooltip=['Sales', 'count()']
    ).transform_filter(selection).transform_aggregate(
        count='count()', groupby=['Sales']
    ).transform_window(
        rank='rank(count)', sort=[alt.SortField('count', order='descending')]
    ).transform_filter(alt.datum.rank <= 10).properties(title='Top 10 Sales', height=300)

    # Layout
    top_row = (chart_main | chart_donut).resolve_scale(color='independent')
    st.altair_chart(top_row, use_container_width=True)
    st.altair_chart(chart_sales, use_container_width=True)

with tab2:
    st.markdown("### บทสรุปสำหรับผู้บริหาร")
    st.markdown(f"""
    <div class="insight-box">
    <b>📈 ผลประกอบการ:</b> ภาพรวมธุรกิจ NIPT ในปี 2025 มีการเติบโตอย่างต่อเนื่อง 
    โดยทำกำไรสะสมรวม <b>{total_gain:,.0f} บาท</b> จากจำนวนเคสทั้งหมด <b>{total_samples} เคส</b>
    </div>
    <div class="insight-box">
    <b>🏆 จุดพีคของปี:</b> เดือนที่ทำผลงานได้ดีที่สุดคือ <b>{best_month}</b> 
    ซึ่งสะท้อนถึงความสำเร็จของทีมขายและการตลาดในช่วงดังกล่าว
    </div>
    <div class="insight-box">
    <b>⏱️ ประสิทธิภาพ:</b> ค่าเฉลี่ย TAT อยู่ที่ <b>{avg_tat:.2f} วัน</b> 
    ซึ่งถือว่ารวดเร็วและเป็นจุดแข็งในการแข่งขัน
    </div>
    """, unsafe_allow_html=True)