import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone
import calendar

# ==========================================
# 1. ページ設定 & 日付取得
# ==========================================
st.set_page_config(page_title="Personal Finance", page_icon="💸", layout="centered")

JST = timezone(timedelta(hours=+9), 'JST')
now_jst = datetime.now(JST)
today = now_jst.date()
current_month = today.strftime("%Y-%m")
days_in_month = calendar.monthrange(today.year, today.month)[1]
remaining_days = max(1, days_in_month - today.day + 1)

# ==========================================
# 2. Googleスプレッドシート設定
# ==========================================
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

@st.cache_resource
def get_gspread_client():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def get_sheet():
    client = get_gspread_client()
    # "FinanceApp" という名前のスプレッドシートを使用
    return client.open("FinanceApp")

# ==========================================
# 3. データ処理ヘルパー関数 (GSheets版)
# ==========================================
def get_monthly_data(month):
    sheet = get_sheet().worksheet("monthly_settings")
    records = sheet.get_all_records()
    
    # 今月のデータを探す
    for row in records:
        if str(row.get("month")) == month:
            return dict(row)
            
    # ★今月のデータがない場合、前月のデータを引き継ぐ（コード①の月またぎ対策を復元）
    data = {
        "month": month,
        "bank_balance": 0, "income_job": 0, "income_allowance": 0, "subs": 0, 
        "mercari_bill": 0, "paypay_bill": 0, "mercari_debt": 0, "paypay_debt": 0,
        "job_date": 25, "allowance_date": 1, "subs_date": 1, "mercari_date": 27, "paypay_date": 27,
        "job_paid": 0, "allowance_paid": 0, "mercari_paid": 0, "paypay_paid": 0
    }
    
    if records:
        # monthカラムで降順ソートして最新（前月）のデータを取得
        sorted_records = sorted(records, key=lambda x: str(x.get("month", "")), reverse=True)
        prev_row = sorted_records[0]
        data.update({
            "bank_balance": prev_row.get("bank_balance", 0),
            "income_job": prev_row.get("income_job", 0),
            "income_allowance": prev_row.get("income_allowance", 0),
            "subs": prev_row.get("subs", 0),
            "mercari_debt": prev_row.get("mercari_debt", 0),
            "paypay_debt": prev_row.get("paypay_debt", 0),
            "job_date": prev_row.get("job_date", 25), 
            "allowance_date": prev_row.get("allowance_date", 1),
            "subs_date": prev_row.get("subs_date", 1), 
            "mercari_date": prev_row.get("mercari_date", 27),
            "paypay_date": prev_row.get("paypay_date", 27)
        })
    
    save_monthly_data(month, data)
    return data

def save_monthly_data(month, data):
    sheet = get_sheet().worksheet("monthly_settings")
    records = sheet.get_all_records()
    
    df = pd.DataFrame(records)
    data["month"] = month # 確実に入れる
    
    if df.empty or month not in df["month"].values:
        # 新規月の場合
        new_df = pd.DataFrame([data])
        df = pd.concat([df, new_df], ignore_index=True)
    else:
        # 既存月の上書き
        for key, value in data.items():
            df.loc[df["month"] == month, key] = value
            
    # NaNを0で埋める（型エラー防止）
    df = df.fillna(0)
    
    # シートをクリアして再書き込み
    sheet.clear()
    sheet.update(values=[df.columns.values.tolist()] + df.values.tolist())

def add_daily_expense(target_date, amount, payment_method, category, memo):
    # 履歴への追加
    sheet_daily = get_sheet().worksheet("daily_expenses")
    row_data = [target_date.strftime("%Y-%m-%d"), amount, payment_method, category, memo]
    sheet_daily.append_row(row_data)
    
    # 残高への自動反映
    m_data = get_monthly_data(current_month)
    
    if category == "特別収入":
        # 収入の場合 (口座に入金、またはクレカの負債が減る)
        if payment_method == "現金/口座": m_data["bank_balance"] += amount
        elif payment_method == "メルカリクレカ": m_data["mercari_debt"] = max(0, m_data["mercari_debt"] - amount)
        elif payment_method == "PayPayクレカ": m_data["paypay_debt"] = max(0, m_data["paypay_debt"] - amount)
    else:
        # 通常支出・特別支出の場合
        if payment_method == "現金/口座": m_data["bank_balance"] -= amount
        elif payment_method == "メルカリクレカ": m_data["mercari_debt"] += amount
        elif payment_method == "PayPayクレカ": m_data["paypay_debt"] += amount

    save_monthly_data(current_month, m_data)

def get_monthly_expenses(month):
    sheet_daily = get_sheet().worksheet("daily_expenses")
    records = sheet_daily.get_all_records()
    
    expenses = {"現金/口座": 0, "メルカリクレカ": 0, "PayPayクレカ": 0}
    for row in records:
        if str(row.get("date", "")).startswith(month):
            # 特別収入は支出合計から除外
            if row.get("category") != "特別収入":
                method = row.get("payment_method")
                if method in expenses:
                    expenses[method] += int(row.get("amount", 0))
    return expenses

def toggle_status(column, current_val):
    m_data = get_monthly_data(current_month)
    m_data[column] = 1 if current_val == 0 else 0
    save_monthly_data(current_month, m_data)

# ==========================================
# 4. キャッシュフロー予測ロジック (コード①完全復元)
# ==========================================
m_data = get_monthly_data(current_month)
spent_this_month = get_monthly_expenses(current_month)

# 1日〜月末までの「未完了」イベントを辞書にまとめる
events = {day: 0 for day in range(1, 32)}

# ★チェックボックスがオフ（未完了）のものだけを予測に含める
if m_data['job_paid'] == 0: events[m_data['job_date']] += m_data['income_job']
if m_data['allowance_paid'] == 0: events[m_data['allowance_date']] += m_data['income_allowance']
if m_data['mercari_paid'] == 0: events[m_data['mercari_date']] -= m_data['mercari_bill']
if m_data['paypay_paid'] == 0: events[m_data['paypay_date']] -= m_data['paypay_bill']
# ※サブスク代(subs)はクレカ請求に含まれる前提のため、予測からは引かない

future_income = sum(v for day, v in events.items() if day >= today.day and v > 0)
future_expense = sum(abs(v) for day, v in events.items() if day >= today.day and v < 0)

# 今月これから自由に使える現金 = 現在の口座残高 + これから入る収入 - これから払うクレカ代
available_total = m_data['bank_balance'] + future_income - future_expense
daily_allowance = available_total // remaining_days if available_total > 0 else 0

# 日々の口座残高シミュレーション（資金ショート判定用）
sim_balance = m_data['bank_balance']
shortfall_warnings = []

for d in range(today.day, days_in_month + 1):
    sim_balance += events[d] # その日の収入・支出を反映
    if sim_balance < 0:
        shortfall_warnings.append({"day": d, "shortfall": abs(sim_balance)})
    # 毎日使える額を消費すると仮定して翌日へ
    sim_balance -= daily_allowance

current_funds = m_data['bank_balance'] + future_income
net_worth = current_funds - m_data['mercari_debt'] - m_data['paypay_debt']

# ==========================================
# 5. UI レンダリング (カテゴリ・メモ入力を追加反映)
# ==========================================
st.title("💸 Finance Tracker")

# --- ① 開いてパッと入力（連動機能つき） ---
with st.container():
    st.markdown("### ☕ 食費・買い物の記録")
    st.caption("ここで入力すると、自動で現在の口座残高やクレカ負債に反映されます。")
    with st.form(key="daily_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        expense_date = col1.date_input("日付", value=today)
        payment_method = col2.selectbox("支払方法", ["現金/口座", "メルカリクレカ", "PayPayクレカ"])
        
        col_cat, col_amt = st.columns(2)
        category = col_cat.selectbox("カテゴリ", ["通常支出", "特別支出", "特別収入"])
        expense_amount = col_amt.number_input("金額(円)", min_value=0, step=100)
        
        col_memo, col_submit = st.columns([3, 1])
        memo = col_memo.text_input("メモ (任意)")
        submit = col_submit.form_submit_button("追加")
        
        if submit and expense_amount > 0:
            add_daily_expense(expense_date, expense_amount, payment_method, category, memo)
            # 通知メッセージもカテゴリに合わせてわかりやすく
            action_text = "収入として残高に加算" if category == "特別収入" else "支出として記録"
            st.success(f"¥{expense_amount:,} を {payment_method} ({category}) で{action_text}しました！")
            st.rerun()

st.divider()

# --- ② 毎日の指標 ---
st.markdown("### 📊 今月の状況")
c1, c2 = st.columns(2)
c1.metric(f"今日使える上限 (残り{remaining_days}日)", f"¥{daily_allowance:,}")
c2.metric("月末まで自由に使える現金", f"¥{available_total:,}")

st.caption(f"今月の利用額計: 現金(¥{spent_this_month['現金/口座']:,}) / メルカリ(¥{spent_this_month['メルカリクレカ']:,}) / PayPay(¥{spent_this_month['PayPayクレカ']:,})")

# --- ③ 支払い判定とアラート ---
st.markdown("### 🚨 支払い判定")
if available_total < 0:
    shortfall = abs(available_total)
    st.error(f"⚠️ **最終的な資金不足**\n\n今月末までにトータルで **¥{shortfall:,}** 不足します。")
    st.warning(f"💡 **対策**: 支払日までに口座へ入金するか、今月のメルカリまたはPayPay請求のうち、**最低 ¥{shortfall:,} 以上を分割払い(リボ)に変更**してください。")
elif shortfall_warnings:
    first_shortfall = shortfall_warnings[0]
    st.error(f"⚠️ **一時的な資金ショート予測！**\n\nトータルでは足りますが、**{first_shortfall['day']}日の支払日時点** で口座残高が **¥{first_shortfall['shortfall']:,}** マイナスになります。給料日とのタイミングのズレにご注意ください。")
elif available_total < 5000:
    st.warning(f"⚠️ 支払いは間に合いますが、今月の残り資金が **¥{available_total:,}** とギリギリです。")
else:
    st.success("💳 クレカの引き落としはスケジュール通り問題なく行えます。")

st.info(f"ℹ️ **サブスク代 (¥{m_data['subs']:,})** はクレカの請求額に含まれている前提のため、二重引き落としを防ぐ目的で口座残高予測からは引いていません。")

st.divider()

# --- ④ ステータス管理 ---
st.markdown("### ✅ 収入・支払いの完了チェック")
st.caption("実際の口座で引き落としや振込が完了したらチェックを入れてください。予測から除外され計算が正確になります。")

col_chk1, col_chk2 = st.columns(2)
if col_chk1.checkbox(f"バイト代受取 (¥{m_data['income_job']:,})", value=bool(m_data['job_paid'])):
    if m_data['job_paid'] == 0: toggle_status('job_paid', 0); st.rerun()
else:
    if m_data['job_paid'] == 1: toggle_status('job_paid', 1); st.rerun()

if col_chk1.checkbox(f"小遣い受取 (¥{m_data['income_allowance']:,})", value=bool(m_data['allowance_paid'])):
    if m_data['allowance_paid'] == 0: toggle_status('allowance_paid', 0); st.rerun()
else:
    if m_data['allowance_paid'] == 1: toggle_status('allowance_paid', 1); st.rerun()

if col_chk2.checkbox(f"メルカリ支払済 (¥{m_data['mercari_bill']:,})", value=bool(m_data['mercari_paid'])):
    if m_data['mercari_paid'] == 0: toggle_status('mercari_paid', 0); st.rerun()
else:
    if m_data['mercari_paid'] == 1: toggle_status('mercari_paid', 1); st.rerun()

if col_chk2.checkbox(f"PayPay支払済 (¥{m_data['paypay_bill']:,})", value=bool(m_data['paypay_paid'])):
    if m_data['paypay_paid'] == 0: toggle_status('paypay_paid', 0); st.rerun()
else:
    if m_data['paypay_paid'] == 1: toggle_status('paypay_paid', 1); st.rerun()

st.divider()

# --- ⑤ 財政予想・残高確認 ---
st.markdown("### 🔮 財政予想 & クレカ総利用残高")
c3, c4 = st.columns(2)
c3.metric("メルカリ総利用残高", f"¥{m_data['mercari_debt']:,}", f"今月請求: ¥{m_data['mercari_bill']:,}", delta_color="inverse")
c4.metric("PayPay総利用残高", f"¥{m_data['paypay_debt']:,}", f"今月請求: ¥{m_data['paypay_bill']:,}", delta_color="inverse")

if net_worth >= 0:
    st.success(f"💡 **現在の実質純資産: ¥{net_worth:,}**\n\n(すべてのクレカ残高を一括返済しても手元にお金が残ります)")
else:
    st.error(f"📉 **現在の実質純資産: ¥{net_worth:,}**\n\n(全クレカ残高を引くとマイナスです。将来の収入で返す必要があります)")

st.divider()

# --- ⑥ 月次データ設定 ---
with st.expander("⚙️ 基本データの設定・修正 (月1回・ズレた時)"):
    with st.form("monthly_data_form"):
        st.caption("※日々の記録を入力すると自動で残高が変わるため、ここは実際の口座額とズレが生じた時の修正や、月初の予定入力に使用してください。")
        
        st.markdown("**💰 口座残高・収入予定**")
        new_balance = st.number_input("銀行口座 現在の実際の残高", value=m_data['bank_balance'], step=1000)
        
        col_in1, col_in2 = st.columns([1, 2])
        new_job_date = col_in1.number_input("バイト代(日)", 1, 31, m_data['job_date'])
        new_job = col_in2.number_input("バイト代(金額)", value=m_data['income_job'], step=1000)
        
        col_in3, col_in4 = st.columns([1, 2])
        new_allowance_date = col_in3.number_input("小遣い(日)", 1, 31, m_data['allowance_date'])
        new_allowance = col_in4.number_input("小遣い(金額)", value=m_data['income_allowance'], step=1000)
        
        st.markdown("**💳 固定費・今月の請求予定**")
        col_s1, col_s2 = st.columns([1, 2])
        new_subs_date = col_s1.number_input("サブスク(日)", 1, 31, m_data['subs_date'])
        new_subs = col_s2.number_input("サブスク(合計額・把握用)", value=m_data['subs'], step=100)
        
        col_m1, col_m2 = st.columns([1, 2])
        new_mercari_date = col_m1.number_input("メルカリ(日)", 1, 31, m_data['mercari_date'])
        new_mercari_bill = col_m2.number_input("メルカリ(請求額)", value=m_data['mercari_bill'], step=1000)
        
        col_p1, col_p2 = st.columns([1, 2])
        new_paypay_date = col_p1.number_input("PayPay(日)", 1, 31, m_data['paypay_date'])
        new_paypay_bill = col_p2.number_input("PayPay(請求額)", value=m_data['paypay_bill'], step=1000)

        st.markdown("**📉 クレカ総利用残高 (未払い負債の合計)**")
        col_debt1, col_debt2 = st.columns(2)
        new_mercari_debt = col_debt1.number_input("メルカリ 総残高", value=m_data['mercari_debt'], step=1000)
        new_paypay_debt = col_debt2.number_input("PayPay 総残高", value=m_data['paypay_debt'], step=1000)
        
        if st.form_submit_button("保存して更新"):
            m_data.update({
                "bank_balance": new_balance,
                "income_job": new_job, "job_date": new_job_date,
                "income_allowance": new_allowance, "allowance_date": new_allowance_date,
                "subs": new_subs, "subs_date": new_subs_date,
                "mercari_bill": new_mercari_bill, "mercari_date": new_mercari_date,
                "paypay_bill": new_paypay_bill, "paypay_date": new_paypay_date,
                "mercari_debt": new_mercari_debt,
                "paypay_debt": new_paypay_debt
            })
            save_monthly_data(current_month, m_data)
            st.success("更新しました！")
            st.rerun()