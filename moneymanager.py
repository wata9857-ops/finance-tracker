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

# 以下の3行を新しく追加してください
@st.cache_resource
def get_worksheet(sheet_name):
    return get_sheet().worksheet(sheet_name)
def get_sheet():
    client = get_gspread_client()
    # "FinanceApp" という名前のスプレッドシートを使用
    return client.open("FinanceApp")

# ==========================================
# 3. データ処理ヘルパー関数 (GSheets版)
# ==========================================
def get_monthly_data(month):
    sheet = get_worksheet("monthly_settings")
    # get_all_records() ではなく get_all_values() を使い、ヘッダー破損によるデータ切り捨てを防ぐ
    all_values = sheet.get_all_values()
    
    correct_keys = [
        "month", "bank_balance", "income_job", "income_allowance", "subs", 
        "mercari_bill", "paypay_bill", "mercari_debt", "paypay_debt",
        "job_date", "allowance_date", "subs_date", "mercari_date", "paypay_date",
        "job_paid", "allowance_paid", "mercari_paid", "paypay_paid"
    ]
    
    records = []
    if len(all_values) > 0:
        # 1行目が '202'（日付）から始まらない場合は、誤ったヘッダー行とみなしてスキップ
        start_idx = 1 if not str(all_values[0][0]).startswith("202") else 0
        for row in all_values[start_idx:]:
            # 足りない列は0で埋める
            padded_row = row + [0] * (len(correct_keys) - len(row))
            parsed_row = []
            for idx, val in enumerate(padded_row[:len(correct_keys)]):
                if idx == 0:
                    parsed_row.append(str(val)) # monthは文字列として保持
                else:
                    try:
                        parsed_row.append(int(float(val)) if str(val).strip() != "" else 0)
                    except:
                        parsed_row.append(0)
            records.append(dict(zip(correct_keys, parsed_row)))
            
    # 今月のデータを探す
    for row in records:
        if str(row.get("month")) == month:
            return dict(row)
            
    # ★今月のデータがない場合、前月のデータを引き継ぐ
    data = {
        "month": month,
        "bank_balance": 0, "income_job": 0, "income_allowance": 0, "subs": 0, 
        "mercari_bill": 0, "paypay_bill": 0, "mercari_debt": 0, "paypay_debt": 0,
        "job_date": 25, "allowance_date": 1, "subs_date": 1, "mercari_date": 27, "paypay_date": 27,
        "job_paid": 0, "allowance_paid": 0, "mercari_paid": 0, "paypay_paid": 0
    }
    
    if records:
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
    sheet = get_worksheet("monthly_settings")
    all_values = sheet.get_all_values()
    
    correct_keys = [
        "month", "bank_balance", "income_job", "income_allowance", "subs", 
        "mercari_bill", "paypay_bill", "mercari_debt", "paypay_debt",
        "job_date", "allowance_date", "subs_date", "mercari_date", "paypay_date",
        "job_paid", "allowance_paid", "mercari_paid", "paypay_paid"
    ]
    
    records = []
    if len(all_values) > 0:
        start_idx = 1 if not str(all_values[0][0]).startswith("202") else 0
        for row in all_values[start_idx:]:
            padded_row = row + [0] * (len(correct_keys) - len(row))
            parsed_row = []
            for idx, val in enumerate(padded_row[:len(correct_keys)]):
                if idx == 0:
                    parsed_row.append(str(val))
                else:
                    try:
                        parsed_row.append(int(float(val)) if str(val).strip() != "" else 0)
                    except:
                        parsed_row.append(0)
            records.append(dict(zip(correct_keys, parsed_row)))
            
    df = pd.DataFrame(records, columns=correct_keys)
    data["month"] = month
    
    if df.empty or month not in df["month"].values:
        new_df = pd.DataFrame([data])
        df = pd.concat([df, new_df], ignore_index=True)
    else:
        for key, value in data.items():
            df.loc[df["month"] == month, key] = value
            
    df = df.fillna(0)
    sheet.clear()
    sheet.update(values=[df.columns.values.tolist()] + df.values.tolist())

def add_daily_expense(target_date, amount, payment_method, category, memo):
    # 履歴への追加
    sheet_daily = get_worksheet("daily_expenses")
    row_data = [target_date.strftime("%Y-%m-%d"), amount, payment_method, category, memo]
    sheet_daily.append_row(row_data)
    
    # 残高への自動反映
    m_data = get_monthly_data(current_month)
    
    if category == "特別収入":
        # 収入の場合
        if payment_method == "現金/口座": m_data["bank_balance"] += amount
        elif payment_method == "メルカリクレカ": m_data["mercari_debt"] = max(0, m_data["mercari_debt"] - amount)
        elif payment_method == "PayPayクレカ": m_data["paypay_debt"] = max(0, m_data["paypay_debt"] - amount)
    elif category == "クレカ先払い(繰り上げ返済)":
        # 支払方法で対象のクレカが選ばれている場合のみ処理（現金のままだとお金が消滅するため）
        if payment_method in ["メルカリクレカ", "PayPayクレカ"]:
            m_data["bank_balance"] -= amount
            if payment_method == "メルカリクレカ": 
                m_data["mercari_debt"] = max(0, m_data["mercari_debt"] - amount)
            elif payment_method == "PayPayクレカ": 
                m_data["paypay_debt"] = max(0, m_data["paypay_debt"] - amount)
    else:
        # 通常支出・特別支出の場合
        if payment_method == "現金/口座": m_data["bank_balance"] -= amount
        elif payment_method == "メルカリクレカ": m_data["mercari_debt"] += amount
        elif payment_method == "PayPayクレカ": m_data["paypay_debt"] += amount

    save_monthly_data(current_month, m_data)

def get_monthly_expenses(month):
    sheet_daily = get_worksheet("daily_expenses")
    daily_values = sheet_daily.get_all_values()
    
    expenses = {"現金/口座": 0, "メルカリクレカ": 0, "PayPayクレカ": 0}
    if len(daily_values) == 0:
        return expenses
        
    start_idx = 1 if not str(daily_values[0][0]).startswith("202") else 0
    
    for row in daily_values[start_idx:]:
        padded = row + [""] * (5 - len(row))
        if str(padded[0]).startswith(month):
            # 特別収入とクレカ先払い（資金移動）は純粋な消費ではないため除外
            if padded[3] not in ["特別収入", "クレカ先払い(繰り上げ返済)"]:
                method = padded[2]
                if method in expenses:
                    try:
                        expenses[method] += int(float(padded[1]) if str(padded[1]).strip() else 0)
                    except:
                        pass
    return expenses

def toggle_status(column, current_val, amount=0, type="income", debt_col=None):
    m_data = get_monthly_data(current_month)
    
    # 0 -> 1 (未完了から完了へ：チェックを入れた時)
    if current_val == 0:
        m_data[column] = 1
        if type == "income":
            m_data["bank_balance"] += amount # 収入なら残高を増やす
        elif type == "expense":
            m_data["bank_balance"] -= amount # 支出なら残高を減らす
            if debt_col:
                m_data[debt_col] = max(0, m_data[debt_col] - amount) # クレカなら未払い総額も減らす
                
    # 1 -> 0 (完了から未完了へ：チェックを外した時（間違えた時の取り消し機能）)
    else:
        m_data[column] = 0
        if type == "income":
            m_data["bank_balance"] -= amount # 収入の取り消し
        elif type == "expense":
            m_data["bank_balance"] += amount # 支出の取り消し
            if debt_col:
                m_data[debt_col] += amount # 減らした未払い総額を元に戻す
                
    save_monthly_data(current_month, m_data)

# ==========================================
# 4. キャッシュフロー予測ロジック
# ==========================================
m_data = get_monthly_data(current_month)
spent_this_month = get_monthly_expenses(current_month)

# 全月データを取得 (5ヶ月先の予測 & 後続のUI表示用)
sheet_m = get_worksheet("monthly_settings")
all_values_m = sheet_m.get_all_values()
correct_keys = [
    "month", "bank_balance", "income_job", "income_allowance", "subs", 
    "mercari_bill", "paypay_bill", "mercari_debt", "paypay_debt",
    "job_date", "allowance_date", "subs_date", "mercari_date", "paypay_date",
    "job_paid", "allowance_paid", "mercari_paid", "paypay_paid"
]
all_m_records_list = []
if len(all_values_m) > 0:
    start_idx = 1 if not str(all_values_m[0][0]).startswith("202") else 0
    for row in all_values_m[start_idx:]:
        padded_row = row + [0] * (len(correct_keys) - len(row))
        all_m_records_list.append(dict(zip(correct_keys, padded_row)))

all_m_records = {str(r.get("month")): r for r in all_m_records_list if str(r.get("month"))}

# 約5ヶ月(150日)先までのイベントを構築
long_term_events = {}
for offset in range(6): 
    y = today.year + (today.month + offset - 1) // 12
    m = (today.month + offset - 1) % 12 + 1
    m_str = f"{y}-{m:02d}"
    
    # 該当月のデータがない場合は今月のデータをベースに仮作成
    m_dict = all_m_records.get(m_str, m_data.copy())
    
    # 未来の月はすべて未払い(0)としてシミュレーションする
    is_current = (offset == 0)
    job_paid = int(float(m_dict.get('job_paid', 0))) if is_current else 0
    allowance_paid = int(float(m_dict.get('allowance_paid', 0))) if is_current else 0
    mercari_paid = int(float(m_dict.get('mercari_paid', 0))) if is_current else 0
    paypay_paid = int(float(m_dict.get('paypay_paid', 0))) if is_current else 0

    def add_event(date_val, name, amount):
        try:
            d = int(float(date_val))
            max_d = calendar.monthrange(y, m)[1]
            safe_d = max(1, min(d, max_d))
            event_date = datetime(y, m, safe_d).date()
            if event_date >= today:
                if event_date not in long_term_events: long_term_events[event_date] = []
                long_term_events[event_date].append({"name": name, "amount": amount})
        except: pass

    if job_paid == 0: add_event(m_dict.get('job_date', 25), "バイト代", int(float(m_dict.get('income_job', 0))))
    if allowance_paid == 0: add_event(m_dict.get('allowance_date', 1), "小遣い", int(float(m_dict.get('income_allowance', 0))))
    if mercari_paid == 0: add_event(m_dict.get('mercari_date', 27), "メルカード引落", -int(float(m_dict.get('mercari_bill', 0))))
    if paypay_paid == 0: add_event(m_dict.get('paypay_date', 27), "PayPay引落", -int(float(m_dict.get('paypay_bill', 0))))

# 150日先までのシミュレーション実行
sim_balance = m_data['bank_balance']
long_term_chart_data = []
short_term_chart_data = [] # 今月用のグラフデータ
shortfall_warnings = []
simulation_log = [] 
bankruptcy_date = None
min_future_balance = sim_balance

end_date = today + timedelta(days=150)
curr_date = today

future_income_1m = 0
future_expense_1m = 0

while curr_date <= end_date:
    daily_events = long_term_events.get(curr_date, [])
    if daily_events:
        daily_in = sum(ev["amount"] for ev in daily_events if ev["amount"] > 0)
        daily_out = sum(abs(ev["amount"]) for ev in daily_events if ev["amount"] < 0)
        sim_balance += (daily_in - daily_out)
        
        # 今月分のログはUI互換のために残す
        if curr_date.month == today.month:
            future_income_1m += daily_in
            future_expense_1m += daily_out
            simulation_log.append({
                "日付": f"{curr_date.month}/{curr_date.day}",
                "イベント": ", ".join([ev["name"] for ev in daily_events]),
                "入出金": daily_in - daily_out,
                "予想口座残高": sim_balance
            })
    
    # グラフ用データの記録
    long_term_chart_data.append({"日付": curr_date.strftime("%Y-%m-%d"), "残高": sim_balance})
    if curr_date.month == today.month:
        short_term_chart_data.append({"日付": f"{curr_date.month}/{curr_date.day}", "残高": sim_balance})
    
    # 最も残高が少なくなる底の金額（真の余力）を更新
    if sim_balance < min_future_balance:
        min_future_balance = sim_balance
        
    # 初めてマイナスになった日（破産日）を記録
    if sim_balance < 0 and bankruptcy_date is None:
        bankruptcy_date = curr_date
        
    # 今月の一時的なショート警告
    if sim_balance < 0 and curr_date.month == today.month and not any(w['day'] == curr_date.day for w in shortfall_warnings):
        shortfall_warnings.append({"day": curr_date.day, "shortfall": abs(sim_balance)})

    curr_date += timedelta(days=1)

# ==========================================
# 真の「使える金額」の再検証
# ==========================================
# ① 表面上の今月の余力
available_total_1_month = m_data['bank_balance'] + future_income_1m - future_expense_1m

# ② 5ヶ月先までの一番厳しい残高（これが真の余力）
true_available_total = max(0, min_future_balance)

# 将来の赤字を避けるため、1ヶ月の余力と長期の余力のうち「厳しい方」を採用
available_total = min(available_total_1_month, true_available_total)
daily_allowance = available_total // remaining_days if available_total > 0 else 0

current_funds = m_data['bank_balance'] + future_income_1m
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
        # ★カテゴリに「クレカ先払い(繰り上げ返済)」を追加
        category = col_cat.selectbox("カテゴリ", ["通常支出", "特別支出", "特別収入", "クレカ先払い(繰り上げ返済)"])
        expense_amount = col_amt.number_input("金額(円)", min_value=0, step=100)
        
        col_memo, col_submit = st.columns([3, 1])
        memo = col_memo.text_input("メモ (任意)")
        submit = col_submit.form_submit_button("追加")
        
        if submit and expense_amount > 0:
            add_daily_expense(expense_date, expense_amount, payment_method, category, memo)
            # ★通知メッセージも操作に合わせてわかりやすく変更
            if category == "特別収入":
                action_text = "収入として残高に加算"
            elif category == "クレカ先払い(繰り上げ返済)":
                action_text = "先払いとして処理（口座と負債からマイナス）"
            else:
                action_text = "支出として記録"
            st.success(f"¥{expense_amount:,} を {payment_method} ({category}) で{action_text}しました！")
            st.rerun()

st.divider()

# --- ② 毎日の指標 ---
st.markdown("### 📊 今月の状況")
c1, c2 = st.columns(2)
c1.metric(f"今日使える上限 (残り{remaining_days}日)", f"¥{daily_allowance:,}")
c2.metric("月末まで自由に使える現金", f"¥{available_total:,}")

st.caption(f"今月の利用額計: 現金(¥{spent_this_month['現金/口座']:,}) / メルカリ(¥{spent_this_month['メルカリクレカ']:,}) / PayPay(¥{spent_this_month['PayPayクレカ']:,})")

# --- ③ 支払い判定とアラート (高度な数学的救済ロジック搭載) ---
# 150日間の予測データから、一番最初に残高がマイナスになる日（破綻日）と不足額を取得
first_fail_record = next((row for row in long_term_chart_data if row["残高"] < 0), None)

if first_fail_record:
    fail_date_str = first_fail_record["日付"]
    fail_amount = abs(first_fail_record["残高"])
    
    if fail_date_str.startswith(f"{today.year}-{today.month:02d}"):
        st.error(f"🚨 **【資金ショート確定】** 今月の **{fail_date_str}** の時点で **¥{fail_amount:,}** が不足し、口座残高がパンクします。")
    else:
        st.error(f"🚨 **【将来の資金ショート確定】** このままいくと **{fail_date_str}** に資金が底をつき、**¥{fail_amount:,}** の赤字になります。")

    with st.expander("🆘 数学的根拠に基づく高度な救済シミュレーション", expanded=True):
        st.markdown("将来のすべての収入・支出予測データ（150日分）の配列を再計算し、破綻を回避するための具体的な数値を導き出しました。")
        
        # --- 解決策A: 日々の支出削減（自力での解決） ---
        st.markdown(f"#### 💡 プランA: 破綻日までの支出ペースを下げる")
        fail_date_obj = datetime.strptime(fail_date_str, "%Y-%m-%d").date()
        days_to_fail = (fail_date_obj - today).days
        
        if days_to_fail > 0:
            daily_cut = fail_amount // days_to_fail
            st.write(f"今日から {fail_date_str} までの **残り{days_to_fail}日間**、1日あたりの変動費（食費や交際費）を現在のペースより **¥{daily_cut:,}** ずつ減らしてください。")
            st.caption("※これに成功すれば、無駄な手数料のかかる分割払いを利用することなく自力で乗り切れます。")
        else:
            st.write("破綻日が本日または過去のため、日々の節約による解決は時間切れです。")

        # --- 解決策B: 最小限の分割払いへの変更 ---
        st.markdown("#### 💳 プランB: クレカ請求の分割払い変更（最適解）")
        total_bill_this_month = int(float(m_data['mercari_bill'])) + int(float(m_data['paypay_bill']))
        
        if total_bill_this_month > 0:
            best_n = None
            best_monthly_payment = 0
            
            # 各クレジットカードの支払日（早い方に合わせる）
            mercari_d = int(float(m_data.get('mercari_date', 27)))
            paypay_d = int(float(m_data.get('paypay_date', 27)))
            payment_date = min(mercari_d if mercari_d > 0 else 27, paypay_d if paypay_d > 0 else 27)
            
            # N回払いのシミュレーション (年利約15% -> 月利1.25%で概算)
            for N in [2, 3, 5, 6, 10, 12, 24]:
                monthly_fee = int(total_bill_this_month * 0.0125)
                monthly_payment = (total_bill_this_month // N) + monthly_fee
                
                is_safe = True
                months_passed = 0
                current_m = today.month
                
                # 未来150日の残高配列に、分割払い変更の「差分（Δ）」を適用して再評価
                for row in long_term_chart_data:
                    d_obj = datetime.strptime(row["日付"], "%Y-%m-%d").date()
                    original_balance = row["残高"]
                    
                    if d_obj.month != current_m:
                        current_m = d_obj.month
                        months_passed += 1
                        
                    # 支払日を何回跨いだか計算
                    payment_count = months_passed + (1 if d_obj.day >= payment_date else 0)
                    
                    if payment_count > 0:
                        # 分割にしたことで「今月払うはずだった全額」が浮き、代わりに「N回分の分割費用」が引かれる
                        # ただし、分割回数(N)の上限を超えた支払いは発生しない
                        actual_payments = min(payment_count, N)
                        adjusted_balance = original_balance + total_bill_this_month - (monthly_payment * actual_payments)
                    else:
                        adjusted_balance = original_balance
                        
                    if adjusted_balance < 0:
                        is_safe = False
                        break  # この分割回数では未来で破綻するためループを抜けて次のNを試す
                        
                if is_safe:
                    best_n = N
                    best_monthly_payment = monthly_payment
                    break  # 最小の安全な分割回数が見つかった時点で確定
            
            if best_n:
                st.success(f"**【数学的最適解】クレカ請求のうち ¥{total_bill_this_month:,} を「{best_n}回払い」に変更してください。**")
                st.write(f"この操作により、今月の支払いは **¥{best_monthly_payment:,}** に下がり、5ヶ月先までシミュレーションしても口座残高はショートしません。")
                total_fee = int(total_bill_this_month * 0.0125 * best_n)
                st.caption(f"※年利15%の手数料を考慮した計算です。トータルで約 ¥{total_fee:,} の手数料を余分に払うことになります。生活に余裕が出た月に一括繰り上げ返済することを推奨します。")
            else:
                st.error("❌ **【打つ手なし】最大24回の分割払いに変更しても、将来どこかで必ずマイナスになります。**")
                st.write("計算上、借金をして支払いを先送りしても根本的な赤字は解消しません。**シフトを増やして収入を上げるか、不用品（使っていない撮影機材や参考書など）を売却して即金を作る**以外に数学的な解決策はありません。")
        else:
            st.write("今月のクレカ請求がないため、分割払いへの変更による救済は利用できません。")

elif shortfall_warnings:
    st.warning("⚠️ **一時的な資金ショート予測！**\n\n月末トータルでは足りますが、給料日などの前に引き落としが来て一時的に残高がマイナスになります。")
    for w in shortfall_warnings:
        st.warning(f"👉 **{today.month}/{w['day']}** の時点で口座残高が **¥{w['shortfall']:,}** 足りなくなります！事前に口座へ入金してください。")

elif available_total < 5000:
    st.warning(f"⚠️ 支払いは間に合いますが、本当に使える残り資金が **¥{available_total:,}** とギリギリです。")
else:
    st.success("💳 固定費・クレカの引き落としはスケジュール通り問題なく行える見込みです。")

# 再検証メッセージの表示
if true_available_total < available_total_1_month:
    st.warning(f"🔍 **「今月使える金額」の再検証結果**\n\n今月だけで見ると ¥{available_total_1_month:,} 余るように見えますが、数ヶ月先の請求を考慮すると赤字になるため、**本当に今月使っていい上限額は ¥{available_total:,} に下方修正** されました。")

st.info(f"ℹ️ **サブスク代 (¥{m_data['subs']:,})** はクレカの請求額に含まれている前提のため、二重引き落としを防ぐ目的で口座残高予測からは引いていません。")

# 残高推移グラフを追加（短期・長期の2段構成）
st.markdown("#### 📊 短期予測グラフ (今月末まで)")
if short_term_chart_data:
    df_chart_short = pd.DataFrame(short_term_chart_data).set_index("日付")
    chart_color_short = "#ff4b4b" if shortfall_warnings or available_total_1_month < 0 else "#00d46a"
    st.line_chart(df_chart_short, y="残高", color=chart_color_short)

st.markdown("#### 📈 長期予測グラフ (5ヶ月先まで)")
if long_term_chart_data:
    df_chart_long = pd.DataFrame(long_term_chart_data).set_index("日付")
    # 長期グラフは色を変えて視認性を高める（破産なら赤、安全なら青系）
    chart_color_long = "#ff4b4b" if bankruptcy_date else "#3282b8"
    st.line_chart(df_chart_long, y="残高", color=chart_color_long)

st.divider()

# カラム横並びを廃止し、タブで横幅を広く確保して見やすさを向上
tab1, tab2 = st.tabs(["🔮 今月の未完了・支出シミュレーション", "📜 過去の支出・収入ログ"])

with tab1:
    if simulation_log:
        df_sim = pd.DataFrame(simulation_log)
        # column_configを使用してリッチなテーブル表示
        st.dataframe(
            df_sim,
            use_container_width=True,
            hide_index=True,
            column_config={
                "日付": st.column_config.TextColumn("日付", width="small"),
                "イベント": st.column_config.TextColumn("イベント", width="medium"),
                "入出金": st.column_config.NumberColumn("入出金", format="¥%d"),
                "予想口座残高": st.column_config.NumberColumn("予想口座残高", format="¥%d")
            }
        )
    else:
        st.info("今月の未完了の収入・支出予定はありません。")

with tab2:
    sheet_daily = get_worksheet("daily_expenses")
    daily_values = sheet_daily.get_all_values()
    
    all_expenses = []
    if len(daily_values) > 0:
        daily_keys = ['date', 'amount', 'payment_method', 'category', 'memo']
        start_idx = 1 if not str(daily_values[0][0]).startswith("202") else 0
        for row in daily_values[start_idx:]:
            padded = row + [""] * (5 - len(row))
            all_expenses.append(dict(zip(daily_keys, padded)))
            
    if all_expenses:
        df_expenses = pd.DataFrame(all_expenses)
        
        # 文字列結合ではなく、数値として保持してPandas Styleで色付けする
        def clean_amount(val):
            try:
                return int(float(val))
            except:
                return 0
                
        df_expenses['amount_num'] = df_expenses['amount'].apply(clean_amount)
        df_expenses['実金額'] = df_expenses.apply(lambda r: r['amount_num'] if r['category'] == '特別収入' else -r['amount_num'], axis=1)
        
        # カテゴリごとにアイコンを付与して視認性アップ
        def get_cat_icon(cat):
            icons = {"通常支出": "🛒", "特別支出": "💸", "特別収入": "💰", "クレカ先払い(繰り上げ返済)": "💳"}
            return f"{icons.get(cat, '📌')} {cat}"
            
        df_expenses['カテゴリ'] = df_expenses['category'].apply(get_cat_icon)
        
        df_display = df_expenses[['date', 'カテゴリ', 'payment_method', '実金額', 'memo']].rename(
            columns={'date': '日付', 'payment_method': '支払方法', 'memo': 'メモ'}
        )
        df_display = df_display.sort_values(by='日付', ascending=False).reset_index(drop=True)
        
        # マイナスは赤、プラスは緑にスタイル適用
        def color_amounts(val):
            if isinstance(val, (int, float)):
                color = '#ff4b4b' if val < 0 else '#00d46a'
                return f'color: {color}; font-weight: bold;'
            return ''
            
        styled_df = df_display.style.map(color_amounts, subset=['実金額']).format({"実金額": "¥{:,.0f}"})
        
        st.dataframe(styled_df, use_container_width=True, hide_index=True)
    else:
        st.write("履歴がありません。")

# --- ④ ステータス管理 ---
st.markdown("### ✅ 収入・支払いの完了チェック")
st.caption("実際の口座で引き落としや振込が完了したらチェックを入れてください。自動で口座残高や未払い総額が計算・保存されます。（間違えて押しても、チェックを外せば元に戻ります！）")

col_chk1, col_chk2 = st.columns(2)

# バイト代
is_job_paid = bool(m_data['job_paid'])
if col_chk1.checkbox(f"バイト代受取 (¥{m_data['income_job']:,})", value=is_job_paid):
    if not is_job_paid: 
        toggle_status('job_paid', 0, amount=m_data['income_job'], type="income")
        st.rerun()
else:
    if is_job_paid: 
        toggle_status('job_paid', 1, amount=m_data['income_job'], type="income")
        st.rerun()

# 小遣い
is_allowance_paid = bool(m_data['allowance_paid'])
if col_chk1.checkbox(f"小遣い受取 (¥{m_data['income_allowance']:,})", value=is_allowance_paid):
    if not is_allowance_paid: 
        toggle_status('allowance_paid', 0, amount=m_data['income_allowance'], type="income")
        st.rerun()
else:
    if is_allowance_paid: 
        toggle_status('allowance_paid', 1, amount=m_data['income_allowance'], type="income")
        st.rerun()

# メルカード
is_mercari_paid = bool(m_data['mercari_paid'])
if col_chk2.checkbox(f"メルカード支払済 (¥{m_data['mercari_bill']:,})", value=is_mercari_paid):
    if not is_mercari_paid: 
        toggle_status('mercari_paid', 0, amount=m_data['mercari_bill'], type="expense", debt_col='mercari_debt')
        st.rerun()
else:
    if is_mercari_paid: 
        toggle_status('mercari_paid', 1, amount=m_data['mercari_bill'], type="expense", debt_col='mercari_debt')
        st.rerun()

# PayPayカード
is_paypay_paid = bool(m_data['paypay_paid'])
if col_chk2.checkbox(f"PayPayカード支払済 (¥{m_data['paypay_bill']:,})", value=is_paypay_paid):
    if not is_paypay_paid: 
        toggle_status('paypay_paid', 0, amount=m_data['paypay_bill'], type="expense", debt_col='paypay_debt')
        st.rerun()
else:
    if is_paypay_paid: 
        toggle_status('paypay_paid', 1, amount=m_data['paypay_bill'], type="expense", debt_col='paypay_debt')
        st.rerun()

st.divider()

# --- ⑤ 財政予想・クレカ利用枠確認 ---
st.markdown("### 🔮 クレカ利用状況 & 財政予想")
st.caption("※メルペイやPayPayの「電子マネー残高（チャージ分）」は、現金と同じように『銀行口座』の残高に含めて管理するとわかりやすいです！")

# ★来月以降に登録されている分割払い等の請求額をスキャンして合計する
# （※「4. キャッシュフロー」で取得済みの all_m_records_list を再利用しAPI通信を削減）
future_mercari_bills = 0
future_paypay_bills = 0

for row in all_m_records_list:
    row_month = str(row.get("month", "")).strip()
    if not row_month:
        continue # 空のデータはスキップ
        
    # 【修正点1】今月以降のデータで、かつ未払いのものを足し合わせる (>= に変更)
    if row_month >= current_month:
        try:
            # 【修正点2】CSVから取得した値は文字列のため、必ず数値に変換してから「0(未払い)かどうか」を判定する
            is_mercari_paid = int(float(row.get("mercari_paid", 0) or 0))
            is_paypay_paid = int(float(row.get("paypay_paid", 0) or 0))
            
            if is_mercari_paid == 0:
                future_mercari_bills += int(float(row.get("mercari_bill", 0) or 0))
            if is_paypay_paid == 0:
                future_paypay_bills += int(float(row.get("paypay_bill", 0) or 0))
        except ValueError:
            pass # 数値に変換できない不正なデータが入っていた場合はスキップ

c3, c4 = st.columns(2)

c3, c4 = st.columns(2)

# 現在の未払い総額 ＋ 未来の分割払い請求額 ＝ 本当の利用残高
# （※ここでスキャンした未来の請求額 future_bills を足し合わせます）
real_mercari_debt = m_data['mercari_debt'] + future_mercari_bills
real_paypay_debt = m_data['paypay_debt'] + future_paypay_bills

mercari_avail = 100000 - real_mercari_debt
paypay_avail = 30000 - real_paypay_debt

c3.metric("メルカード 利用可能枠 (上限10万)", f"¥{mercari_avail:,}", f"未払い総額: -¥{real_mercari_debt:,}", delta_color="normal")
c4.metric("PayPayカード 利用可能枠 (上限3万)", f"¥{paypay_avail:,}", f"未払い総額: -¥{real_paypay_debt:,}", delta_color="normal")

# 純資産も、未来の分割払いを含めた「本当の負債」で計算
real_net_worth = current_funds - real_mercari_debt - real_paypay_debt

if real_net_worth >= 0:
    st.success(f"💡 **現在の実質純資産: ¥{real_net_worth:,}**\n\n(すべてのクレカ未払い分を一括返済しても手元にお金が残ります)")
else:
    st.error(f"📉 **現在の実質純資産: ¥{real_net_worth:,}**\n\n(全クレカ未払い分を引くとマイナスです。将来の収入で返す必要があります)")

st.divider()

# --- ⑥ 月次データ設定 ---
with st.expander("⚙️ 基本データの設定・修正 (月1回・ズレた時・未来の請求予定)"):
    
    # --- 追加機能：編集する月を選ぶ ---
    month_options = []
    for i in range(12): # 今月から12ヶ月先まで選択可能にする
        m = today.month + i
        y = today.year + (m - 1) // 12
        m = (m - 1) % 12 + 1
        month_options.append(f"{y}-{m:02d}")
        
    edit_target_month = st.selectbox("📝 予定を入力・編集する月を選択", month_options)
    edit_m_data = get_monthly_data(edit_target_month)
    
    # 選択された月の「日付」オブジェクトを安全に作る関数（カレンダー表示用）
    def get_date_for_month(day_int):
        y, m = map(int, edit_target_month.split("-"))
        max_day = calendar.monthrange(y, m)[1]
        safe_d = max(1, min(int(day_int), max_day)) # 存在しない日（2月30日など）を防止
        return datetime(y, m, safe_d).date()
        
    with st.form("monthly_data_form"):
        st.caption(f"※現在【 {edit_target_month} 】のデータを編集しています。カレンダーから日付を選んでください。")
        
        st.markdown("**💰 口座・電子マネー残高 & 収入予定**")
        new_balance = st.number_input("銀行口座＋チャージ済電子マネーの合計残高", value=edit_m_data['bank_balance'], step=1000)
        
        col_in1, col_in2 = st.columns([1, 2])
        new_job_date_obj = col_in1.date_input("バイト代(日付)", value=get_date_for_month(edit_m_data['job_date']))
        new_job = col_in2.number_input("バイト代(金額)", value=edit_m_data['income_job'], step=1000)
        
        col_in3, col_in4 = st.columns([1, 2])
        new_allowance_date_obj = col_in3.date_input("小遣い(日付)", value=get_date_for_month(edit_m_data['allowance_date']))
        new_allowance = col_in4.number_input("小遣い(金額)", value=edit_m_data['income_allowance'], step=1000)
        
        st.markdown("**💳 固定費・今月の請求予定**")
        col_s1, col_s2 = st.columns([1, 2])
        new_subs_date_obj = col_s1.date_input("サブスク(日付)", value=get_date_for_month(edit_m_data['subs_date']))
        new_subs = col_s2.number_input("サブスク(合計額・把握用)", value=edit_m_data['subs'], step=100)
        
        col_m1, col_m2 = st.columns([1, 2])
        new_mercari_date_obj = col_m1.date_input("メルカード(支払日)", value=get_date_for_month(edit_m_data['mercari_date']))
        new_mercari_bill = col_m2.number_input("メルカード(今月の請求額)", value=edit_m_data['mercari_bill'], step=1000)
        
        col_p1, col_p2 = st.columns([1, 2])
        new_paypay_date_obj = col_p1.date_input("PayPayカード(支払日)", value=get_date_for_month(edit_m_data['paypay_date']))
        new_paypay_bill = col_p2.number_input("PayPayカード(今月の請求額)", value=edit_m_data['paypay_bill'], step=1000)

        st.markdown("**📉 クレカの未払い利用総額 (まだ口座から引き落とされていない利用分の合計)**")
        st.caption("※すでに支払いが確定している「今月の請求額」と、分割などで「来月以降に支払う額」の合計です。")
        col_debt1, col_debt2 = st.columns(2)
        new_mercari_debt = col_debt1.number_input("メルカード 未払い総額", value=edit_m_data['mercari_debt'], step=1000)
        new_paypay_debt = col_debt2.number_input("PayPayカード 未払い総額", value=edit_m_data['paypay_debt'], step=1000)
        
        if st.form_submit_button("保存して更新"):
            edit_m_data.update({
                "bank_balance": new_balance,
                "income_job": new_job, "job_date": new_job_date_obj.day, # カレンダーから「日」だけを抽出して保存
                "income_allowance": new_allowance, "allowance_date": new_allowance_date_obj.day,
                "subs": new_subs, "subs_date": new_subs_date_obj.day,
                "mercari_bill": new_mercari_bill, "mercari_date": new_mercari_date_obj.day,
                "paypay_bill": new_paypay_bill, "paypay_date": new_paypay_date_obj.day,
                "mercari_debt": new_mercari_debt,
                "paypay_debt": new_paypay_debt
            })
            save_monthly_data(edit_target_month, edit_m_data)
            st.success(f"✅ {edit_target_month} のデータを更新しました！")
            st.rerun()