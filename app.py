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
    return df_summary, df_detail, df_capital, df_rent

try:
    df_summary, df_detail, df_capital, df_rent = load_data()
except Exception as e:
    st.error(f"讀取 Excel 檔案發生錯誤: {e}")
    st.stop()

# ==================== 🔒 密碼登入驗證保護 ====================
# 您可以隨時在此修改您的存取密碼
CORRECT_PASSWORD = "14789"

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
            
    st.stop() # 尚未登入前，停止載入後續的儀表板畫面

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
    
    p_spent = 0.0
    b_spent = 519605.48
    c_spent = 3444000.00
    total_spent = p_spent + b_spent + c_spent
    total_budget = 85447167.00
    total_remaining = total_budget - total_spent
    total_progress = (total_spent / total_budget * 100) if total_budget > 0 else 0

    # 產生自訂 HTML 表格（完美靠右對齊數字與會計格式）
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
        text-align: right; /* 第二欄開始全部靠右對齊 */
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
        <td>NT$ 0</td>
        <td>NT$ 48,124,541</td>
        <td>0.00%</td>
      </tr>
      <tr>
        <td>業務費與租金 (經常門)</td>
        <td>NT$ 19,051,087.48</td>
        <td>NT$ 519,605.48</td>
        <td>NT$ 18,531,482</td>
        <td>2.73%</td>
      </tr>
      <tr>
        <td>資本門</td>
        <td>NT$ 18,271,539</td>
        <td>NT$ 3,444,000</td>
        <td>NT$ 14,827,539</td>
        <td>18.85%</td>
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
            '預算金額': [48124541.00, 19051087.48, 18271539.00],
            '已支用/核銷金額': [0.0, 519605.48, 3444000.00]
        })
        fig_bar = px.bar(door_summary_bar, x='經費門類', y=['預算金額', '已支用/核銷金額'], barmode='group', title="三大經費門類預算與支用對比")
        st.plotly_chart(fig_bar, use_container_width=True)

# 2. 經常門統計 (含人事費跟業務費)
elif menu == "💵 經常門統計 (含人事費與業務費)":
    st.header("💵 經常門經費統計 (含人事費與業務費/租金)")
    st.markdown("依據支用標準，經常門包含**人事費**與**業務費**（租金歸類於業務費下）。")
    
    st.subheader("👤 一、人事費支用與預算標準 (總預算: NT$ 48,124,541)")
    personnel_data = [
        ["碩士級研究助理", "-", "-", "-", "NT$ 484,811", "研究助理到職前之經費預算，可釋出 245,145 元供總計畫經費運用。"],
        ["獎金", "-", "-", "-", "NT$ 32,506,430", "留任獎金及醫學中心職涯發展獎金：護理師、醫檢師、放射師等每季獎勵在職留任人員。"],
        ["獎金", "-", "-", "-", "NT$ 15,133,300", "計畫執行人員獎勵（個人獎勵與科室獎勵）。"]
    ]
    personnel_df = pd.DataFrame(personnel_data, columns=['項目', '單價', '單位', '數量', '合計 (NTD)', '支用說明或編列基準'])
    st.dataframe(personnel_df, use_container_width=True)
    
    st.markdown("---")
    
    st.subheader("📝 二、業務費與租金編列標準明細 (總預算: NT$ 19,051,087)")
    business_data = [
        ["租金", "NT$ 47,789", "台/月", "285.5", "NT$ 13,642,989", "目標3-AI機器人引進：機器人使用租金(AI照護型)*50台*5.7個月"],
        ["租金", "NT$ 349,167", "月", "5.0", "NT$ 1,757,098", "目標2-人工智慧模型：手術室排程預測與管理系統租金"],
        ["租金", "NT$ 685,714", "月", "4.6", "NT$ 3,163,132", "目標1-建置員工智管家App：員工智管家App租金"],
        ["租金", "NT$ 56,667", "月", "4", "NT$ 226,668", "目標2-人工智慧模型：急診室AI檢傷分類戰情室租金"],
        ["餐費", "NT$ 140", "人", "-", "NT$ 130,000", "範疇一每月2次會議，各範疇會議30人*140元*2次*15個月"],
        ["雜支費", "-", "-", "-", "NT$ 20,000", "執行本計畫相關雜項支出"],
        ["文具紙張", "-", "批", "1", "NT$ 7,000", "執行計畫所需油墨、碳粉匣、紙張、文具等費用"],
        ["印刷", "-", "批", "1", "NT$ 10,000", "執行計畫所需書表影印、印裝費"],
        ["出席費", "NT$ 2,500", "人", "36", "NT$ 90,000", "評選會外部委員出席審查費用 2,500元/人/次"],
        ["國內旅費", "NT$ 1,400", "-", "3", "NT$ 4,200", "相關人員及出席專家之國內旅費（高鐵/台鐵費）"]
    ]
    business_df = pd.DataFrame(business_data, columns=['項目', '單價', '單位', '數量', '合計 (NTD)', '支用說明或編列基準'])
    st.dataframe(business_df, use_container_width=True)
    
    st.markdown("---")
    
    st.subheader("📊 三、業務費各細項實際執行狀況與餘額統計")
    actual_spent = all_details.groupby(cat_col)[total_amt_col].sum().to_dict()
    
    biz_stats = []
    budget_dict = {'餐費': 130000, '雜支費': 20000, '文具紙張': 7000, '印刷': 10000, '出席費': 90000, '國內旅費': 4200, '租金': 18789974}
    
    for item_name, b_amt in budget_dict.items():
        spent = actual_spent.get(item_name, 0.0)
        remaining = b_amt - spent
        progress = (spent / b_amt * 100) if b_amt > 0 else 0
        biz_stats.append({
            '項目': item_name,
            '預算金額 (A)': f"NT$ {b_amt:,.0f}",
            '已核銷支出 (B)': f"NT$ {spent:,.2f}" if spent % 1 != 0 else f"NT$ {spent:,.0f}",
            '剩餘可用餘額 (C)': f"NT$ {remaining:,.2f}" if remaining % 1 != 0 else f"NT$ {remaining:,.0f}",
            '執行進度 (%)': f"{progress:.2f}%"
        })
    
    df_biz_stats = pd.DataFrame(biz_stats)
    st.dataframe(df_biz_stats, use_container_width=True)
    
    st.markdown("### 🏢 業務費 - 租金細項執行清單")
    st.dataframe(df_rent, use_container_width=True)
    
    st.info(f"💡 **已沖帳額度統計（E欄總和）**：目前可沖帳額度總計為 **NT$ {total_chongzhang_quota:,.0f}** 元。")

# 3. 資本門統計
elif menu == "🏢 資本門統計":
    st.header("🏢 資本門經費統計與設備採購追蹤")
    st.markdown("依據支用標準，資本門編列以百分之三十為上限，主要用於設備購置與裝置。")
    
    cap_df = df_capital.iloc[1:5, [0, 1, 2, 3, 4, 5, 6, 7, 8]].copy()
    cap_df.columns = ['發票開立日期', '發票號碼', '品名', '數量', '單價', '總計(含稅)', '核銷單列印日期', '核銷案號', '說明/備註']
    st.dataframe(cap_df, use_container_width=True)
    
    st.metric("資本門已核銷總額", "NT$ 3,444,000")

# 4. 明細表
elif menu == "📋 明細表":
    st.header("📋 經費支用詳細明細與待沖帳管理")
    
    tab1, tab2, tab3 = st.tabs(["📄 全部流水帳", "⏳ 待沖帳項目清單", "🔍 可沖帳額度明細"])
    
    with tab1:
        st.dataframe(all_details, use_container_width=True)
            
    with tab2:
        pending_df = all_details[all_details[cat_col].isna() | (all_details[cat_col] == '待沖帳')].copy()
        st.dataframe(pending_df, use_container_width=True)
        st.warning(f"目前共有 {len(pending_df)} 筆待沖帳項目，總計核銷總額 NT$ {pending_df[total_amt_col].sum():,.0f} 元。")
        
    with tab3:
        chongzhang_df = all_details[all_details[chongzhang_amt_col] > 0].copy()
        st.dataframe(chongzhang_df, use_container_width=True)
        st.success(f"有設定可沖帳額度的項目共 {len(chongzhang_df)} 筆，可沖帳額度總計 NT$ {chongzhang_df[chongzhang_amt_col].sum():,.0f} 元。")

# 5. 新增記帳
elif menu == "➕ 新增記帳":
    st.header("➕ 記錄新的經費花費")
    st.markdown("完全對齊明細表 A-K 欄架構：可填寫發票日期、會計核銷科目、實際品項、核銷總額、可沖帳額度、實際支出、代墊人員與核銷案號。")
    
    with st.form("expense_form"):
        col1, col2 = st.columns(2)
        with col1:
            exp_date = st.date_input("發票/收據日期 (A欄)", value=date.today())
            exp_category = st.selectbox("會計核銷科目 (B欄)", ["餐費", "文具紙張", "印刷", "雜支費", "出席費", "國內旅費", "租金", "資本門", "待沖帳", "代墊還款"])
            exp_item = st.text_input("實際購買品項 (C欄)", value="")
            exp_total_amt = st.number_input("核銷總額 / 發票總額 (D欄)", min_value=0, value=1400, step=100)
        with col2:
            exp_chongzhang = st.number_input("可沖帳額度 (E欄)", min_value=0, value=0, step=50)
            exp_actual = st.number_input("實際支出 (F欄)", min_value=0, value=1400, step=50)
            exp_payer = st.text_input("代墊 / 借支人員", value="")
            exp_status = st.selectbox("還款狀態", ["V", "未還款", "-"])
            exp_case_no = st.text_input("核銷案號 (J欄)", value="")
            
        submitted = st.form_submit_button("💾 儲存這筆花費紀錄")
        if submitted:
            new_row = {
                "發票/收據日期": str(exp_date),
                cat_col: exp_category,
                item_col: exp_item,
                total_amt_col: exp_total_amt,
                chongzhang_amt_col: exp_chongzhang,
                actual_amt_col: exp_actual,
                "代墊 / 借支人員": exp_payer,
                "自科室借款": 0,
                "還款狀態": exp_status,
                "核銷單\n列印日期": "",
                "核銷案號\n(黃底未借支)": exp_case_no
            }
            st.session_state['new_expenses'].append(new_row)
            st.success(f"✅ 成功新增記錄！科目：**{exp_category}**，核銷總額：**${exp_total_amt:,.0f}** 元（可沖帳額度：**${exp_chongzhang:,.0f}** 元）。")

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>範疇一經費管理系統面板 | Powered by Streamlit & Pandas</p>", unsafe_allow_html=True)