import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="ERP Cash Collection System", layout="wide")
st.title("💵 Daily Cash Collection & Weekly Settlement System")

# 1. Database Files Setup
cash_file = "cash_collection.csv"
if not os.path.exists(cash_file):
    df_cash_init = pd.DataFrame([
        {"Date": "2026-09-01", "Collector": "Ko Kyaw", "Customer": "City Mart", "Amount": 150000},
        {"Date": "2026-09-03", "Collector": "Aung Aung", "Customer": "Ocean Store", "Amount": 200000},
        {"Date": "2026-09-08", "Collector": "Ko Kyaw", "Customer": "Grand Hotel", "Amount": 180000}
    ])
    df_cash_init.to_csv(cash_file, index=False)

df_cash = pd.read_csv(cash_file)
df_cash["Date"] = pd.to_datetime(df_cash["Date"])

# 2. Sidebar - Daily Cash Entry Form (With Customer Name)
st.sidebar.header("📥 Daily Cash Entry (Cashier)")
with st.sidebar.form("cash_entry_form"):
    entry_date = st.date_input("Collection Date")
    collector_name = st.text_input("Collector Name")
    customer_name = st.text_input("Customer Name")
    amount = st.number_input("Amount Collected (MMK)", min_value=0, step=1000)
    cash_submitted = st.form_submit_button("Save Cash Record")
    
    if cash_submitted and collector_name and customer_name and amount > 0:
        new_cash = pd.DataFrame([{
            "Date": str(entry_date), 
            "Collector": collector_name, 
            "Customer": customer_name,
            "Amount": amount
        }])
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
