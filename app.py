import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# 設定網頁版面
st.set_page_config(
    page_title="範疇一經費監控與管理儀表板",
    page_icon="💰",
    layout="wide"
)

FILE_DETAIL = '114~115年度_範疇一經費支用紀錄明細.xlsx'
FILE_BUDGET = '★範疇一經費115.08(8.27經服中心公佈).xlsx'

@st.cache_data
def load_data():
    df_summary = pd.read_excel(FILE_DETAIL, sheet_name='業務費總表')
    df_detail = pd.read_excel(FILE_DETAIL, sheet_name='明細表 ')
    df_capital = pd.read_excel(FILE_DETAIL, sheet_name='資本門')
    df_rent = pd.read_excel(FILE_DETAIL, sheet_name='業務費-租金')

    try:
        df_personnel_biz = pd.read_excel(FILE_DETAIL, sheet_name='業務費-人事費')
    except Exception:
        df_personnel_biz = pd.DataFrame()

    return df_summary, df_detail, df_capital, df_rent, df_personnel_biz

try:
    df_summary, df_detail, df_capital, df_rent, df_personnel_biz = load_data()
except Exception as e:
    st.error(f"讀取 Excel 檔案發生錯誤: {e}")
    st.stop()

# ==================== 🔒 密碼登入驗證保護 ====================
CORRECT_PASSWORD = "14379"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 範疇一經費管理系統 - 請先登入")
    st.markdown("本系統包含計畫經費機敏資料，請輸入存取密碼以繼續。")

    password_input = st.text_input("請輸入系統訪問密碼：", type="password")

    if st.button("登入系統"):
        if password_input == CORRECT_PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ 密碼錯誤，請重新輸入。")

    st.stop()

# ==================== 初始化 Session State 用於即時記帳 ====================
if 'new_expenses' not in st.session_state:
    st.session_state['new_expenses'] = []

# ==================== 標題與介紹 ====================
st.title("📊 範疇一計畫經費監控與管理儀表板")
st.markdown("依據 **115.08 經費分配表**、**經費支用標準** 與最新 **114~115年度經費支用紀錄明細** 即時掌握經費動態。")
st.markdown("---")

# 欄位精準對應明細表
cat_col = '會計核銷科目'
item_col = '實際購買品項'
total_amt_col = '核銷總額'
chongzhang_amt_col = '可沖帳額度'
actual_amt_col = '實際支出'

# 動態計算數據
all_details = pd.concat([df_detail, pd.DataFrame(st.session_state['new_expenses'])], ignore_index=True) if st.session_state['new_expenses'] else df_detail

excel_pending_sum = all_details[all_details[cat_col].isna() | (all_details[cat_col] == '待沖帳')][total_amt_col].sum()
total_pending = excel_pending_sum

# 自動計算 E 欄「可沖帳額度」總和
total_chongzhang_quota = all_details[chongzhang_amt_col].sum()

# 計算剩餘未沖帳金額
remaining_unreimbursed = total_pending - total_chongzhang_quota

# 計算研究助理已核銷總額
ra_budget = 484811.0
if not df_personnel_biz.empty:
    col_name_total = df_personnel_biz.columns[2]
    valid_ra = df_personnel_biz.iloc[1:9].copy()
    ra_spent = pd.to_numeric(valid_ra[col_name_total], errors='coerce').sum()
else:
    ra_spent = 258443.0

ra_remaining = ra_budget - ra_spent
ra_progress = (ra_spent / ra_budget * 100) if ra_budget > 0 else 0

# 動態計算資本門各項目已核銷金額（從 df_capital 抓取）
cap_df_raw = df_capital.iloc[1:12, [1, 2, 5]].copy()
cap_df_raw.columns = ['發票號碼', '品名', '總計']
cap_df_raw['總計'] = pd.to_numeric(cap_df_raw['總計'], errors='coerce').fillna(0)

ai_spent = cap_df_raw[cap_df_raw['品名'].astype(str).str.contains('AI排班', na=False)]['總計'].sum()
loc_spent = cap_df_raw[cap_df_raw['品名'].astype(str).str.contains('定位系統|定位點班', na=False)]['總計'].sum()
ilisten_spent = cap_df_raw[cap_df_raw['品名'].astype(str).str.contains('iListen|數位化溝通平台', na=False)]['總計'].sum()
scanner_spent = cap_df_raw[cap_df_raw['品名'].astype(str).str.contains('掃描機', na=False)]['總計'].sum()
scale_spent = cap_df_raw[cap_df_raw['品名'].astype(str).str.contains('磅秤', na=False)]['總計'].sum()
erp_spent = cap_df_raw[cap_df_raw['品名'].astype(str).str.contains('ERP', na=False)]['總計'].sum()

total_cap_spent = cap_df_raw['總計'].sum()

# 動態計算租金各項目已核銷金額（改抓 F 欄「核銷金額」）
rent_df_raw = df_rent.iloc[1:12, [1, 2, 5, 6, 7]].copy()
rent_df_raw.columns = ['發票號碼', '品名', '核銷金額', '核銷單列印日期', '核銷案號']
rent_df_raw['核銷金額'] = pd.to_numeric(rent_df_raw['核銷金額'], errors='coerce').fillna(0)

robot_rent_spent = rent_df_raw[rent_df_raw['品名'].astype(str).str.contains('機器人', na=False)]['核銷金額'].sum()
surgery_rent_spent = rent_df_raw[rent_df_raw['品名'].astype(str).str.contains('手術智慧排程|手術室排程', na=False)]['核銷金額'].sum()
app_rent_spent = rent_df_raw[rent_df_raw['品名'].astype(str).str.contains('智管家', na=False)]['核銷金額'].sum()
er_rent_spent = rent_df_raw[rent_df_raw['品名'].astype(str).str.contains('急診', na=False)]['核銷金額'].sum()

total_rent_spent = robot_rent_spent + surgery_rent_spent + app_rent_spent + er_rent_spent

# ==================== 頂部核心經費看板 ====================
st.subheader("📌 核心經費總覽看板")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("範疇一總經費上限", "NT$ 85,447,167")
with col2:
    st.metric("人事費總預算 (經常門)", "NT$ 48,124,541", "穩定執行中")
with col3:
    st.metric("業務費總預算 (含租金)", "NT$ 19,051,087")
with col4:
    st.metric("資本門總預算", "NT$ 18,271,539")

st.markdown("")

# 獨立隔開的沖帳追蹤看板
st.subheader("📌 經費沖帳進度追蹤看板")
col5, col6, col7 = st.columns(3)
with col5:
    st.metric("⏳ 待沖帳項目總額", f"NT$ {total_pending:,.0f}")
with col6:
    st.metric("🔍 已沖帳總額 (E欄加總)", f"NT$ {total_chongzhang_quota:,.0f}")
with col7:
    st.metric("📊 剩餘未沖帳金額", f"NT$ {remaining_unreimbursed:,.0f}", delta_color="inverse")

st.markdown("---")

# ==================== 側邊欄導航 ====================
st.sidebar.header("🔍 系統導航")
menu = st.sidebar.radio("選擇功能模組", [
    "📌 總體經費摘要與支用比例",
    "💵 經常門統計 (含人事費與業務費)",
    "🏢 資本門統計",
    "📋 明細表",
    "➕ 新增記帳"
])

# 1. 總體經費摘要與支用比例
if menu == "📌 總體經費摘要與支用比例":
    st.header("📌 總體經費執行摘要與支用比例")

    p_spent = ra_spent
    b_spent = 130000 + 20000 + 7000 + 10000 + 90000 + 4200 + total_rent_spent
    c_spent = total_cap_spent
    total_spent = p_spent + b_spent + c_spent
    total_budget = 85447167.00
    total_remaining = total_budget - total_spent
    total_progress = (total_spent / total_budget * 100) if total_budget > 0 else 0

    html_table = f"""
    <style>
    .accounting-table {{
        width: 100%;
        border-collapse: collapse;
        font-family: sans-serif;
        font-size: 15px;
        background-color: white;
        color: #31333F;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        border-radius: 8px;
        overflow: hidden;
    }}
    .accounting-table th, .accounting-table td {{
        padding: 12px 16px;
        border-bottom: 1px solid #e0e0e0;
    }}
    .accounting-table th {{
        background-color: #f0f2f6;
        font-weight: 600;
        text-align: left;
    }}
    .accounting-table th:nth-child(n+2), .accounting-table td:nth-child(n+2) {{
        text-align: right;
    }}
    .accounting-table tr:hover {{
        background-color: #f8f9fa;
    }}
    .total-row {{
        font-weight: bold;
        background-color: #eef2f7 !important;
    }}
    </style>

    <table class="accounting-table">
      <tr>
        <th>經費門類</th>
        <th>預算金額</th>
        <th>已支用/核銷金額</th>
        <th>剩餘可用金額</th>
        <th>支用比例(%)</th>
      </tr>
      <tr>
        <td>人事費 (經常門)</td>
        <td>NT$ 48,124,541</td>
        <td>NT$ {p_spent:,.2f}</td>
        <td>NT$ {48124541 - p_spent:,.2f}</td>
        <td>{(p_spent / 48124541 * 100):.2f}%</td>
      </tr>
      <tr>
        <td>業務費與租金 (經常門)</td>
        <td>NT$ 19,051,087.48</td>
        <td>NT$ {b_spent:,.2f}</td>
        <td>NT$ {19051087.48 - b_spent:,.2f}</td>
        <td>{(b_spent / 19051087.48 * 100):.2f}%</td>
      </tr>
      <tr>
        <td>資本門</td>
        <td>NT$ 18,271,539</td>
        <td>NT$ {c_spent:,.2f}</td>
        <td>NT$ {18271539 - c_spent:,.2f}</td>
        <td>{(c_spent / 18271539 * 100):.2f}%</td>
      </tr>
      <tr class="total-row">
        <td>📌 範疇一總經費合計</td>
        <td>NT$ {total_budget:,.2f}</td>
        <td>NT$ {total_spent:,.2f}</td>
        <td>NT$ {total_remaining:,.2f}</td>
        <td>{total_progress:.2f}%</td>
      </tr>
    </table>
    """
    st.markdown(html_table, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        door_summary_pie = pd.DataFrame({
            '經費門類': ['人事費 (經常門)', '業務費與租金 (經常門)', '資本門'],
            '預算金額': [48124541.00, 19051087.48, 18271539.00]
        })
        fig_pie = px.pie(door_summary_pie, names='經費門類', values='預算金額', title="三大經費門類預算分配佔比", hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)
    with col_b:
        door_summary_bar = pd.DataFrame({
            '經費門類': ['人事費 (經常門)', '業務費與租金 (經常門)', '資本門'],
            '預算金額': [48124541.00, 19051087