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

# ==================== 圓環圖建立函式（保留比例數字，游標指過去顯示詳細金額） ====================
def make_donut(title, spent, budget, color_done, height=220, title_size=14, center_size=16, show_legend=False):
    rem = max(budget - spent, 0.0)
    pct = (spent / budget * 100) if budget > 0 else 0.0
    fig = px.pie(
        pd.DataFrame({
            '狀態': ['已核銷進度', '剩餘可用額度'],
            '金額': [spent, rem]
        }),
        names='狀態',
        values='金額',
        hole=0.65,
        color='狀態',
        color_discrete_map={'已核銷進度': color_done, '剩餘可用額度': '#3a3d42'},
        title=title
    )
    fig.update_traces(
        textinfo='percent',
        textfont_size=max(center_size - 3, 11),
        hovertemplate='<b>%{label}</b><br>金額：NT$ %{value:,.0f}<br>比例：%{percent}<extra></extra>'
    )
    fig.add_annotation(
        text=f"<b>{pct:.1f}%</b>",
        x=0.5, y=0.5, font_size=center_size, showarrow=False
    )
    fig.update_layout(
        height=height,
        showlegend=show_legend,
        title=dict(font=dict(size=title_size), x=0.5, xanchor='center'),
        margin=dict(t=42, b=12, l=12, r=12)
    )
    return fig

# ==================== 側邊欄導航 ====================
st.sidebar.header("🔍 系統導航")
menu = st.sidebar.radio("選擇功能模組", [
    "📌 總體經費摘要與支用比例",
    "💵 經常門統計 (含人事費與業務費)",
    "🏢 資本門統計",
    "📋 明細表",
    "➕ 新增記帳"
])

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

# 人事費 3 筆項目年底均列為 100% 核銷（總預算 NT$ 48,124,541）
ra_budget_100 = 484811.0
bonus_retain_budget_100 = 32506430.0
bonus_exec_budget_100 = 15133300.0
personnel_full_budget = ra_budget_100 + bonus_retain_budget_100 + bonus_exec_budget_100
personnel_spent_100 = personnel_full_budget

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

# 動態計算租金各項目已核銷金額（抓 F 欄「核銷金額」）
rent_df_raw = df_rent.iloc[1:12, [1, 2, 5, 6, 7]].copy()
rent_df_raw.columns = ['發票號碼', '品名', '核銷金額', '核銷單列印日期', '核銷案號']
rent_df_raw['核銷金額'] = pd.to_numeric(rent_df_raw['核銷金額'], errors='coerce').fillna(0)

robot_rent_spent = rent_df_raw[rent_df_raw['品名'].astype(str).str.contains('機器人', na=False)]['核銷金額'].sum()
surgery_rent_spent = rent_df_raw[rent_df_raw['品名'].astype(str).str.contains('手術智慧排程|手術室排程', na=False)]['核銷金額'].sum()
app_rent_spent = rent_df_raw[rent_df_raw['品名'].astype(str).str.contains('智管家', na=False)]['核銷金額'].sum()
er_rent_spent = rent_df_raw[rent_df_raw['品名'].astype(str).str.contains('急診', na=False)]['核銷金額'].sum()

total_rent_spent = robot_rent_spent + surgery_rent_spent + app_rent_spent + er_rent_spent

# 1. 總體經費摘要與支用比例（保留完整總覽看板與沖帳追蹤看板）
if menu == "📌 總體經費摘要與支用比例":
    st.title("📊 範疇一計畫經費監控與管理儀表板")
    st.markdown("依據 **115.08 經費分配表**、**經費支用標準** 與最新 **114~115年度經費支用紀錄明細** 即時掌握經費動態。")
    st.markdown("---")

    st.subheader("📌 核心經費總覽看板")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("範疇一總經費上限", "NT$ 85,447,167")
    with col2:
        st.metric("人事費總預算 (經常門)", "NT$ 48,124,541", "年底列為 100% 核銷")
    with col3:
        st.metric("業務費總預算 (含租金)", "NT$ 19,051,087")
    with col4:
        st.metric("資本門總預算", "NT$ 18,271,539")

    st.markdown("")

    st.subheader("📌 經費沖帳進度追蹤看板")
    col5, col6, col7 = st.columns(3)
    with col5:
        st.metric("⏳ 待沖帳項目總額", f"NT$ {total_pending:,.0f}")
    with col6:
        st.metric("🔍 已沖帳總額 (E欄加總)", f"NT$ {total_chongzhang_quota:,.0f}")
    with col7:
        st.metric("📊 剩餘未沖帳金額", f"NT$ {remaining_unreimbursed:,.0f}", delta_color="inverse")

    st.markdown("---")
    st.header("📌 總體經費執行摘要與支用比例")

    p_spent = personnel_spent_100
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
        <td>人事費 (經常門，年底 100% 核銷)</td>
        <td>NT$ 48,124,541</td>
        <td>NT$ {p_spent:,.2f}</td>
        <td>NT$ {48124541 - p_spent:,.2f}</td>
        <td>100.00%</td>
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
            '預算金額': [48124541.00, 19051087.48, 18271539.00],
            '已支用/核銷金額': [p_spent, b_spent, c_spent]
        })
        fig_bar = px.bar(door_summary_bar, x='經費門類', y=['預算金額', '已支用/核銷金額'], barmode='group', title="三大經費門類預算與支用對比")
        st.plotly_chart(fig_bar, use_container_width=True)

# 2. 經常門統計：左邊最大圈圈（經常門總體），右邊小圈圈（人事 3 項 + 租金 4 項）
elif menu == "💵 經常門統計 (含人事費與業務費)":
    st.header("💵 經常門經費統計 (含人事費與業務費/租金)")

    freq_budget = 48124541.0 + 19051087.48
    freq_spent = personnel_spent_100 + (130000 + 20000 + 7000 + 10000 + 90000 + 4200 + total_rent_spent)

    left_col, right_col = st.columns([1.35, 2.65])

    with left_col:
        st.plotly_chart(
            make_donut(
                title="經常門總體核銷",
                spent=freq_spent,
                budget=freq_budget,
                color_done='#2ecc71',
                height=380,
                title_size=18,
                center_size=24,
                show_legend=True
            ),
            use_container_width=True
        )

    with right_col:
        freq_sub_items = [
            ("碩士級研究助理", ra_budget_100, ra_budget_100, '#27ae60'),
            ("留任與職涯獎金", bonus_retain_budget_100, bonus_retain_budget_100, '#27ae60'),
            ("執行人員獎勵", bonus_exec_budget_100, bonus_exec_budget_100, '#27ae60'),
            ("AI機器人租金", robot_rent_spent, 13642989.0, '#1abc9c'),
            ("手術排程租金", surgery_rent_spent, 1757098.0, '#1abc9c'),
            ("智管家App租金", app_rent_spent, 3163132.0, '#1abc9c'),
            ("急診AI戰情租金", er_rent_spent, 226668.0, '#1abc9c'),
        ]
        r1 = st.columns(4)
        for idx in range(4):
            title, sp, bg, col = freq_sub_items[idx]
            with r1[idx]:
                st.plotly_chart(
                    make_donut(title, sp, bg, col, height=180, title_size=13, center_size=14, show_legend=False),
                    use_container_width=True
                )

        r2 = st.columns(4)
        for idx in range(4, 7):
            title, sp, bg, col = freq_sub_items[idx]
            with r2[idx - 4]:
                st.plotly_chart(
                    make_donut(title, sp, bg, col, height=180, title_size=13, center_size=14, show_legend=False),
                    use_container_width=True
                )

    st.markdown("---")

    st.subheader("👤 一、人事費支用與預算標準 (總預算: NT$ 48,124,541｜全數列為 100% 核銷)")
    personnel_data = [
        [
            "碩士級研究助理",
            "NT$ 484,811",
            "NT$ 484,811",
            "100.00%",
            "研究助理到職前之經費預算，可釋出 245,145 元供總計畫經費運用（年底預估全數核銷）。"
        ],
        [
            "獎金",
            "NT$ 32,506,430",
            "NT$ 32,506,430",
            "100.00%",
            "留任獎金及醫學中心職涯發展獎金：護理師、醫檢師、放射師等每季獎勵在職留任人員（年底預估全數核銷）。"
        ],
        [
            "獎金",
            "NT$ 15,133,300",
            "NT$ 15,133,300",
            "100.00%",
            "計畫執行人員獎勵（個人獎勵與科室獎勵，年底預估全數核銷）。"
        ]
    ]
    personnel_df = pd.DataFrame(
        personnel_data,
        columns=['項目', '預算金額 (NTD)', '已核銷金額 (NTD)', '核銷比例', '支用說明或編列基準']
    )
    st.dataframe(personnel_df, use_container_width=True)

    st.markdown("---")
    st.subheader("🏢 三、業務費 - 四大租金專案執行與已核銷月份追蹤")
    rent_tracking_data = [
        ["目標3-AI機器人引進 (AI照護型)", 13642989.0, robot_rent_spent, "7/10~10月", "機器人使用租金 (50台 * 5.7個月)"],
        ["目標2-手術室排程預測與管理系統", 1757098.0, surgery_rent_spent, "7月~10月", "手術室排程預測與管理系統租金"],
        ["目標1-員工智管家 App", 3163132.0, app_rent_spent, "8/17~10月", "員工智管家 App 租金"],
        ["目標2-急診室AI檢傷分類戰情室", 226668.0, er_rent_spent, "尚無", "急診室AI檢傷分類戰情室租金"]
    ]

    rent_stats = []
    for row in rent_tracking_data:
        name, b_val, s_val, months, desc = row
        rem_val = b_val - s_val
        prog = (s_val / b_val * 100) if b_val > 0 else 0
        rent_stats.append({
            '租金專案項目': name,
            '預算金額 (A)': f"NT$ {b_val:,.2f}" if b_val % 1 != 0 else f"NT$ {b_val:,.0f}",
            '已核銷金額 (B)': f"NT$ {s_val:,.2f}" if s_val % 1 != 0 else f"NT$ {s_val:,.0f}",
            '剩餘可用額度 (C)': f"NT$ {rem_val:,.2f}" if rem_val % 1 != 0 else f"NT$ {rem_val:,.0f}",
            '已核銷月份': months,
            '核銷執行進度 (%)': f"{prog:.2f}%",
            '說明': desc
        })
    df_rent_tracking = pd.DataFrame(rent_stats)
    st.dataframe(df_rent_tracking, use_container_width=True)

    st.markdown("---")
    st.subheader("📊 四、業務費各細項實際執行狀況與餘額統計")
    actual_spent = all_details.groupby(cat_col)[total_amt_col].sum().to_dict()

    biz_stats = []
    budget_dict = {'餐費': 130000, '雜支費': 20000, '文具紙張': 7000, '印刷': 10000, '出席費': 90000, '國內旅費': 4200, '租金': 18789974}

    for item_name, b_amt in budget_dict.items():
        spent = actual_spent.get(item_name, 0.0)
        if item_name == '租金':
            spent = total_rent_spent
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
    display_rent_df = df_rent.iloc[1:12, [0, 1, 2, 3, 4, 5, 6, 7, 8]].copy()
    display_rent_df.columns = ['發票開立日期', '發票號碼', '品名', '數量', '單價', '總計(含稅)', '核銷金額(F欄)', '核銷單列印日期', '核銷案號']
    display_rent_df['核銷金額(F欄)'] = display_rent_df['核銷金額(F欄)'].apply(
        lambda x: f"NT$ {x:,.0f}" if pd.notnull(x) and isinstance(x, (int, float)) else x
    )
    st.dataframe(display_rent_df, use_container_width=True)

# 3. 資本門統計：左邊最大圈圈（資本門總體），右邊小圈圈（各系統名稱）
elif menu == "🏢 資本門統計":
    st.header("🏢 資本門經費統計與設備採購追蹤")

    cap_budget = 18271539.0

    left_col, right_col = st.columns([1.35, 2.65])

    with left_col:
        st.plotly_chart(
            make_donut(
                title="資本門總體核銷",
                spent=total_cap_spent,
                budget=cap_budget,
                color_done='#3498db',
                height=380,
                title_size=18,
                center_size=24,
                show_legend=True
            ),
            use_container_width=True
        )

    with right_col:
        cap_sub_items = [
            ("AI排班與加班預測系統", ai_spent, 3570000.0, '#3498db'),
            ("定位點班系統", loc_spent, 1960000.0, '#5dade2'),
            ("iListen員工脈動平台", ilisten_spent, 720000.0, '#2980b9'),
            ("全自動掃描機", scanner_spent, 3281539.0, '#3498db'),
            ("智能磅秤", scale_spent, 740000.0, '#5dade2'),
            ("ERP系統", erp_spent, 8000000.0, '#2980b9'),
        ]
        r1 = st.columns(3)
        for idx in range(3):
            title, sp, bg, col = cap_sub_items[idx]
            with r1[idx]:
                st.plotly_chart(
                    make_donut(title, sp, bg, col, height=180, title_size=13, center_size=14, show_legend=False),
                    use_container_width=True
                )

        r2 = st.columns(3)
        for idx in range(3, 6):
            title, sp, bg, col = cap_sub_items[idx]
            with r2[idx - 3]:
                st.plotly_chart(
                    make_donut(title, sp, bg, col, height=180, title_size=13, center_size=14, show_legend=False),
                    use_container_width=True
                )

    st.markdown("---")
    st.subheader("📋 一、資本門 6 大合併項目核銷進度清單")

    capital_tracking_data = [
        ["AI排班與加班預測系統", 3570000.0, ai_spent, "包含30%簽約金、第一、二階段及第三階段(尾款)"],
        ["定位點班系統", 1960000.0, loc_spent, "加護病房設備室內定位系統建置案（第一期已核銷）"],
        ["iListen員工脈動平台", 720000.0, ilisten_spent, "包含30%簽約與70%尾款（已核銷平台建置費）"],
        ["全自動掃描機", 3281539.0, scanner_spent, "數位病理專案所需（剩餘款由本院自籌）"],
        ["智能磅秤", 740000.0, scale_spent, "擴大科技投資降低工作負荷-智慧量測系統（50個）"],
        ["ERP系統", 8000000.0, erp_spent, "企業資源規劃系統：ERP系統"]
    ]

    cap_stats = []
    for row in capital_tracking_data:
        name, b_val, s_val, desc = row
        rem_val = b_val - s_val
        prog = (s_val / b_val * 100) if b_val > 0 else 0
        cap_stats.append({
            '資本門專案項目': name,
            '預算編列金額 (A)': f"NT$ {b_val:,.2f}" if b_val % 1 != 0 else f"NT$ {b_val:,.0f}",
            '已核銷金額 (B)': f"NT$ {s_val:,.2f}" if s_val % 1 != 0 else f"NT$ {s_val:,.0f}",
            '剩餘可用額度 (C)': f"NT$ {rem_val:,.2f}" if rem_val % 1 != 0 else f"NT$ {rem_val:,.0f}",
            '核銷執行進度 (%)': f"{prog:.2f}%",
            '備註說明': desc
        })
    df_cap_tracking = pd.DataFrame(cap_stats)
    st.dataframe(df_cap_tracking, use_container_width=True)

    st.markdown("---")
    st.subheader("📄 二、資本門原始發票與核銷流水明細")
    cap_df = df_capital.iloc[1:12, [0, 1, 2, 3, 4, 5, 6, 7, 8]].copy()
    cap_df.columns = ['發票開立日期', '發票號碼', '品名', '數量', '單價', '總計(含稅)', '核銷單列印日期', '核銷案號', '說明/備註']
    cap_df['總計(含稅)'] = cap_df['總計(含稅)'].apply(lambda x: f"NT$ {x:,.0f}" if pd.notnull(x) and isinstance(x, (int, float)) else x)
    st.dataframe(cap_df, use_container_width=True)

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
from datetime import datetime
from zoneinfo import ZoneInfo

now_tw = datetime.now(ZoneInfo("Asia/Taipei")).strftime("%Y/%m/%d %H:%M")

st.markdown(
    f"""
    <style>
    .report-timestamp-badge {{
        position: fixed;
        right: 18px;
        bottom: 12px;
        background-color: rgba(30, 32, 37, 0.88);
        color: #e6e8eb;
        border: 1px solid rgba(255, 255, 255, 0.15);
        padding: 5px 11px;
        border-radius: 6px;
        font-size: 12px;
        z-index: 9999;
        pointer-events: none;
    }}
    </style>
    <div class="report-timestamp-badge">更新時間：{now_tw}</div>
    """,
    unsafe_allow_html=True,
)