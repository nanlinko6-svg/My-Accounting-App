import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="ERP Collection & Settlement System", layout="wide")
st.title("💵 Daily Collection & Weekly Settlement System")

# 1. Database Setup
cash_file = "cash_collection.csv"
if not os.path.exists(cash_file):
    df_cash_init = pd.DataFrame([
        {"Date": "2026-09-01", "Collector": "Ko Kyaw", "Customer": "City Mart", "Payment_Type": "Cash", "Cheque_No": "-", "Amount": 150000},
        {"Date": "2026-09-03", "Collector": "Aung Aung", "Customer": "Ocean Store", "Payment_Type": "Cheque", "Cheque_No": "CHQ-998821", "Amount": 200000},
        {"Date": "2026-09-08", "Collector": "Ko Kyaw", "Customer": "Grand Hotel", "Payment_Type": "Cash", "Cheque_No": "-", "Amount": 180000}
    ])
    df_cash_init.to_csv(cash_file, index=False)

df_cash = pd.read_csv(cash_file)
df_cash["Date"] = pd.to_datetime(df_cash["Date"])
df_cash["Cheque_No"] = df_cash["Cheque_No"].fillna("-")

# 2. Sidebar - Data Entry Form (Cash / Cheque)
st.sidebar.header("📥 Daily Collection Entry")
with st.sidebar.form("cash_entry_form"):
    entry_date = st.date_input("Collection Date")
    collector_name = st.text_input("Collector Name")
    customer_name = st.text_input("Customer Name")
    pay_type = st.selectbox("Payment Type", ["Cash", "Cheque"])
    cheque_no = st.text_input("Cheque No. (Optional for Cash)")
    amount = st.number_input("Amount Collected (MMK)", min_value=0, step=1000)
    
    cash_submitted = st.form_submit_button("Save Collection Record")
    
    if cash_submitted and collector_name and customer_name and amount > 0:
        chq_val = cheque_no if (pay_type == "Cheque" and cheque_no) else "-"
        new_cash = pd.DataFrame([{
            "Date": str(entry_date), 
            "Collector": collector_name, 
            "Customer": customer_name,
            "Payment_Type": pay_type,
            "Cheque_No": chq_val,
            "Amount": amount
        }])
        new_cash.to_csv(cash_file, mode="a", header=False, index=False)
        st.success(f"Saved {pay_type} ({amount:,.0f} MMK) from {customer_name}")
        st.rerun()

# 3. Weekly Calculation Logic
df_cash["Week"] = df_cash["Date"].dt.isocalendar().week
df_cash["Year"] = df_cash["Date"].dt.isocalendar().year

# Metrics Overview
total_cash_val = df_cash[df_cash["Payment_Type"] == "Cash"]["Amount"].sum()
total_chq_val = df_cash[df_cash["Payment_Type"] == "Cheque"]["Amount"].sum()
grand_total = df_cash["Amount"].sum()

col1, col2, col3 = st.columns(3)
col1.metric("Grand Total Collected", f"{grand_total:,.0f} MMK")
col2.metric("Total Cash Collected", f"{total_cash_val:,.0f} MMK")
col3.metric("Total Cheque Received", f"{total_chq_val:,.0f} MMK")

st.markdown("---")

# 4. Weekly Summary Table
st.subheader("📅 Weekly Summary Report (သီတင်းပတ်အလိုက် စာရင်းချုပ်)")
weekly_summary = df_cash.groupby(["Year", "Week", "Collector", "Customer", "Payment_Type"])["Amount"].sum().reset_index()
weekly_summary.columns = ["Year", "Week No.", "Collector Name", "Customer Name", "Payment Type", "Total Amount (MMK)"]
st.dataframe(weekly_summary, use_container_width=True)

# 5. Daily Transaction Log Table
st.subheader("📝 Daily Transaction Detail Logs")
st.dataframe(
    df_cash[["Date", "Collector", "Customer", "Payment_Type", "Cheque_No", "Amount"]]
    .sort_values(by="Date", ascending=False), 
    use_container_width=True
)
        new_cash.to_csv(cash_file, mode="a", header=False, index=False)
        st.success(f"Saved {amount:,.0f} MMK from {customer_name} ({collector_name})")
        st.rerun()

# 3. Weekly Calculation Logic
df_cash["Week"] = df_cash["Date"].dt.isocalendar().week
df_cash["Year"] = df_cash["Date"].dt.isocalendar().year

# Metrics Overview
total_cash = df_cash["Amount"].sum()
col1, col2 = st.columns(2)
col1.metric("Total Collected Cash", f"{total_cash:,.0f} MMK")
col2.metric("Total Entry Count", len(df_cash))

st.markdown("---")

# 4. Weekly Summary Table (Collector & Customer Breakdown)
st.subheader("📅 Weekly Summary Report (သီတင်းပတ်အလိုက် ငွေစာရင်းချုပ်)")
weekly_summary = df_cash.groupby(["Year", "Week", "Collector", "Customer"])["Amount"].sum().reset_index()
weekly_summary.columns = ["Year", "Week No.", "Collector Name", "Customer Name", "Total Amount (MMK)"]
st.dataframe(weekly_summary, use_container_width=True)

# 5. Daily Transaction Log Table
st.subheader("📝 Daily Transaction Detail Logs")
st.dataframe(df_cash[["Date", "Collector", "Customer", "Amount"]].sort_values(by="Date", ascending=False), use_container_width=True)
rue)
