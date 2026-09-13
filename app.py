import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="My ERP Accounting", layout="wide")
st.title("📊 Enterprise Accounting & AR Dashboard")

# CSV Database File စစ်ဆေးခြင်း
file_path = "accounting_data.csv"
if not os.path.exists(file_path):
    df_init = pd.DataFrame([
        {"Customer Name": "Aung Aung", "Days Overdue": 15, "Status": "Normal"},
        {"Customer Name": "Kyaw Kyaw", "Days Overdue": 45, "Status": "Warning"},
        {"Customer Name": "Phyu Phyu", "Days Overdue": 75, "Status": "CRITICAL"}
    ])
    df_init.to_csv(file_path, index=False)

# Data ရယူခြင်း
df = pd.read_csv(file_path)

# Sidebar - Data Entry Form
st.sidebar.header("➕ Customer Order Entry")
with st.sidebar.form("entry_form"):
    name = st.text_input("Customer Name")
    days = st.number_input("Days Overdue", min_value=0, step=1)
    submitted = st.form_submit_button("Save Record")
    
    if submitted and name:
        status = "Normal" if days <= 30 else ("Warning" if days <= 60 else "CRITICAL")
        new_data = pd.DataFrame([{"Customer Name": name, "Days Overdue": days, "Status": status}])
        new_data.to_csv(file_path, mode="a", header=False, index=False)
        st.success(f"Saved: {name}")
        st.rerun()

# Dashboard Summary Metrics
col1, col2, col3 = st.columns(3)
col1.metric("Total Customers", len(df))
col2.metric("Critical Overdue (>60 Days)", len(df[df["Status"] == "CRITICAL"]))
col3.metric("Warning Overdue (31-60 Days)", len(df[df["Status"] == "Warning"]))

st.subheader("📋 AR Aging Audit Table")
st.dataframe(df, use_container_width=True)
