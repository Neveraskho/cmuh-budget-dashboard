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

# 動態計算租金各項目已核銷金額（改抓 F 欄「核銷金額」，在 pandas 中為第 6 欄 index 5）
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
    st.