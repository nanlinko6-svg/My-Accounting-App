import streamlit as st
import pandas as pd
import os
import io
from fpdf import FPDF

st.set_page_config(page_title="ERP Collection System", layout="wide")

# 1. Multi-User Credentials List (6 Users)
USERS = {
    "admin": "Admin@2026",
    "finance1": "Fin1@123",
    "finance2": "Fin2@123",
    "finance3": "Fin3@123",
    "finance4": "Fin4@123",
    "finance5": "Fin5@123"
}

# User Authentication System
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""

def login():
    st.title("🔐 System Login")
    with st.form("login_form"):
        username_input = st.text_input("Username").strip().lower()
        password_input = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login")
        
        if submit:
            if username_input in USERS and USERS[username_input] == password_input:
                st.session_state["authenticated"] = True
                st.session_state["username"] = username_input
                st.success("Login Successful!")
                st.rerun()
            else:
                st.error("Invalid Username or Password")

if not st.session_state["authenticated"]:
    login()
    st.stop()

# Logout Button & Active User Info in Sidebar
st.sidebar.markdown(f"👤 **Logged in as:** `{st.session_state['username']}`")
if st.sidebar.button("🔒 Logout"):
    st.session_state["authenticated"] = False
    st.session_state["username"] = ""
    st.rerun()

st.title("💵 Daily Collection & Weekly Settlement System")

# 2. Database Setup
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

# 3. Sidebar - Data Entry Form (Cash / Cheque)
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
        new_cash = pd.DataFrame([{"Date": str(entry_date), "Collector": collector_name, "Customer": customer_name, "Payment_Type": pay_type, "Cheque_No": chq_val, "Amount": amount}])
        new_cash.to_csv(cash_file, mode="a", header=False, index=False)
        st.success(f"Saved {pay_type} ({amount:,.0f} MMK) from {customer_name}")
        st.rerun()

# 4. Weekly Calculation Logic
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

# 5. Export Utility Functions (Excel & A4 PDF)
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Collection_Report')
    return output.getvalue()

def generate_pdf_report(df):
    # A4 Size (210 x 297 mm) Explicit Specification
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Daily Collection & Settlement Report (A4)", ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.ln(5)
    
    # Table Header (Fits A4 width ~ 190mm)
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(25, 8, "Date", 1, 0, 'C')
    pdf.cell(35, 8, "Collector", 1, 0, 'C')
    pdf.cell(45, 8, "Customer", 1, 0, 'C')
    pdf.cell(20, 8, "Type", 1, 0, 'C')
    pdf.cell(30, 8, "Cheque No", 1, 0, 'C')
    pdf.cell(35, 8, "Amount (MMK)", 1, 1, 'C')
    
    # Table Rows
    pdf.set_font("Arial", size=9)
    for _, row in df.iterrows():
        date_str = pd.to_datetime(row['Date']).strftime('%Y-%m-%d')
        pdf.cell(25, 7, str(date_str), 1)
        pdf.cell(35, 7, str(row['Collector'])[:18], 1)
        pdf.cell(45, 7, str(row['Customer'])[:23], 1)
        pdf.cell(20, 7, str(row['Payment_Type']), 1, 0, 'C')
        pdf.cell(30, 7, str(row['Cheque_No']), 1, 0, 'C')
        pdf.cell(35, 7, f"{row['Amount']:,}", 1, 1, 'R')
        
    return bytes(pdf.output(dest='S'))

# Export Buttons Section
st.subheader("📥 Export Reports (Excel / A4 PDF)")
col_ex1, col_ex2 = st.columns(2)

excel_data = convert_df_to_excel(df_cash)
col_ex1.download_button(
    label="📊 Download Excel Report (.xlsx)",
    data=excel_data,
    file_name="Collection_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

pdf_data = generate_pdf_report(df_cash)
col_ex2.download_button(
    label="📄 Download PDF Report (A4 Standard)",
    data=pdf_data,
    file_name="Collection_Report_A4.pdf",
    mime="application/pdf"
)

st.markdown("---")

# 6. Weekly Summary Table
st.subheader("📅 Weekly Summary Report (သီတင်းပတ်အလိုက် စာရင်းချုပ်)")
weekly_summary = df_cash.groupby(["Year", "Week", "Collector", "Customer", "Payment_Type"])["Amount"].sum().reset_index()
weekly_summary.columns = ["Year", "Week No.", "Collector Name", "Customer Name", "Payment Type", "Total Amount (MMK)"]
st.dataframe(weekly_summary, use_container_width=True)

# 7. Daily Transaction Log Table
st.subheader("📝 Daily Transaction Detail Logs")
st.dataframe(df_cash[["Date", "Collector", "Customer", "Payment_Type", "Cheque_No", "Amount"]].sort_values(by="Date", ascending=False), use_container_width=True)
