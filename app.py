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
    df_cash_init = pd.DataFrame(columns=["Date", "Collector", "Customer", "Payment_Type", "Cheque_No", "Amount"])
    df_cash_init.to_csv(cash_file, index=False)

df_cash = pd.read_csv(cash_file)

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

# 4. Sidebar - Delete Specific Row (တစ်ခုချင်းစီ ဖျက်ရန်)
if not df_cash.empty:
    st.sidebar.markdown("---")
    st.sidebar.header("❌ Delete Specific Record")
    
    valid_delete_df = df_cash.dropna(subset=["Collector", "Amount"])
    delete_options = [f"ID {idx}: {row['Date']} | {row['Collector']} | {row['Customer']} ({row['Amount']:,} MMK)" for idx, row in valid_delete_df.iterrows()]
    
    if delete_options:
        selected_option = st.sidebar.selectbox("Select Record to Delete", delete_options)
        if st.sidebar.button("🗑 Delete Selected Record"):
            selected_index = int(selected_option.split(":")[0].replace("ID ", ""))
            df_cash = df_cash.drop(selected_index).reset_index(drop=True)
            df_cash.to_csv(cash_file, index=False)
            st.sidebar.success("Selected record deleted successfully!")
            st.rerun()

# Processing Calculations if data exists
if not df_cash.empty:
    # Clean NaN values
    df_cash = df_cash.dropna(subset=["Amount"]).reset_index(drop=True)
    df_cash["Amount"] = pd.to_numeric(df_cash["Amount"], errors='coerce').fillna(0)
    df_cash["Date_dt"] = pd.to_datetime(df_cash["Date"], errors='coerce')
    df_cash["Cheque_No"] = df_cash["Cheque_No"].fillna("-")
    df_cash["Collector"] = df_cash["Collector"].fillna("-")
    df_cash["Customer"] = df_cash["Customer"].fillna("-")
    df_cash["Payment_Type"] = df_cash["Payment_Type"].fillna("Cash")
    df_cash["Week"] = df_cash["Date_dt"].dt.isocalendar().week
    df_cash["Year"] = df_cash["Date_dt"].dt.isocalendar().year

    # Metrics Overview
    total_cash_val = df_cash[df_cash["Payment_Type"] == "Cash"]["Amount"].sum()
    total_chq_val = df_cash[df_cash["Payment_Type"] == "Cheque"]["Amount"].sum()
    grand_total = df_cash["Amount"].sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("Grand Total Collected", f"{grand_total:,.0f} MMK")
    col2.metric("Total Cash Collected", f"{total_cash_val:,.0f} MMK")
    col3.metric("Total Cheque Received", f"{total_chq_val:,.0f} MMK")

    st.markdown("---")

    # Fixed Excel Export Function (Explicit Headers for All Sheets)
    def convert_df_to_excel(df):
        output = io.BytesIO()
        
        # Prepare Daily Details Data & Rename Headers
        detail_df = df[["Date", "Collector", "Customer", "Payment_Type", "Cheque_No", "Amount"]].sort_values(by="Date", ascending=False).copy()
        detail_df.columns = ["Collection Date", "Collector Name", "Customer Name", "Payment Type", "Cheque No.", "Amount (MMK)"]
        
        # Prepare Weekly Summary Data & Rename Headers
        summary_df = df.groupby(["Year", "Week", "Collector", "Customer", "Payment_Type"])["Amount"].sum().reset_index()
        summary_df.columns = ["Year", "Week No.", "Collector Name", "Customer Name", "Payment Type", "Total Amount (MMK)"]

        # Write to Excel Sheets
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name='Weekly_Summary', header=True)
            detail_df.to_excel(writer, index=False, sheet_name='Daily_Details', header=True)
            
        return output.getvalue()

    # Fixed PDF Generator Function with Summary Totals Section
    def generate_pdf_report(df):
        pdf = FPDF(orientation='P', unit='mm', format='A4')
        pdf.add_page()
        
        # Header Title
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "Daily Collection & Settlement Report (A4)", ln=True, align='C')
        pdf.set_font("Arial", size=10)
        pdf.ln(3)
        
        # Table Headers
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(25, 8, "Date", 1, 0, 'C')
        pdf.cell(35, 8, "Collector", 1, 0, 'C')
        pdf.cell(45, 8, "Customer", 1, 0, 'C')
        pdf.cell(20, 8, "Type", 1, 0, 'C')
        pdf.cell(30, 8, "Cheque No", 1, 0, 'C')
        pdf.cell(35, 8, "Amount (MMK)", 1, 1, 'C')
        
        # Table Data
        pdf.set_font("Arial", size=9)
        for _, row in df.iterrows():
            date_str = str(row['Date'])[:10] if pd.notna(row['Date']) else "-"
            collector_str = str(row['Collector'])[:18]
            customer_str = str(row['Customer'])[:23]
            pay_type_str = str(row['Payment_Type'])
            cheque_str = str(row['Cheque_No'])
            amt_str = f"{row['Amount']:,.0f}"

            pdf.cell(25, 7, date_str, 1, 0, 'C')
            pdf.cell(35, 7, collector_str, 1, 0, 'L')
            pdf.cell(45, 7, customer_str, 1, 0, 'L')
            pdf.cell(20, 7, pay_type_str, 1, 0, 'C')
            pdf.cell(30, 7, cheque_str, 1, 0, 'C')
            pdf.cell(35, 7, amt_str, 1, 1, 'R')
            
        # Summary Section at Bottom of PDF
        pdf.ln(6)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(0, 7, "Collection Summary Totals:", 0, 1, 'L')
        
        c_val = df[df["Payment_Type"] == "Cash"]["Amount"].sum()
        q_val = df[df["Payment_Type"] == "Cheque"]["Amount"].sum()
        g_val = df["Amount"].sum()
        
        pdf.set_font("Arial", size=9)
        pdf.cell(105, 7, "Total Cash Collected:", 1, 0, 'L')
        pdf.cell(85, 7, f"{c_val:,.0f} MMK", 1, 1, 'R')
        
        pdf.cell(105, 7, "Total Cheque Received:", 1, 0, 'L')
        pdf.cell(85, 7, f"{q_val:,.0f} MMK", 1, 1, 'R')
        
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(105, 7, "Grand Total Collected:", 1, 0, 'L')
        pdf.cell(85, 7, f"{g_val:,.0f} MMK", 1, 1, 'R')
            
        # Return Stream Output
        pdf_str = pdf.output(dest='S')
        if isinstance(pdf_str, str):
            return pdf_str.encode('latin-1', errors='replace')
        return bytes(pdf_str)

    # Export Buttons
    st.subheader("📥 Export Reports (Excel Multi-Sheet / A4 PDF)")
    col_ex1, col_ex2 = st.columns(2)

    excel_data = convert_df_to_excel(df_cash)
    col_ex1.download_button(
        label="📊 Download Excel Report (Summary & Detail Sheets)",
        data=excel_data,
        file_name="Collection_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    try:
        pdf_bytes = generate_pdf_report(df_cash)
        col_ex2.download_button(
            label="📄 Download PDF Report (A4 Standard)",
            data=pdf_bytes,
            file_name="Collection_Report_A4.pdf",
            mime="application/pdf"
        )
    except Exception as e:
        col_ex2.error(f"⚠️ PDF အဆင်မပြေပါ: {e}")

    st.markdown("---")

    # 5. Weekly Summary Table
    st.subheader("📅 Weekly Summary Report (သီတင်းပတ်အလိုက် စာရင်းချုပ်)")
    weekly_summary = df_cash.groupby(["Year", "Week", "Collector", "Customer", "Payment_Type"])["Amount"].sum().reset_index()
    weekly_summary.columns = ["Year", "Week No.", "Collector Name", "Customer Name", "Payment Type", "Total Amount (MMK)"]
    st.dataframe(weekly_summary, use_container_width=True)

    # 6. Interactive Editable Daily Logs Table
    st.subheader("📝 Daily Transaction Detail Logs (Editable)")
    st.info("💡 ဇယားထဲတွင် ကလစ်နှိပ်၍ အချက်အလက်များ တိုက်ရိုက်ပြင်ဆင်နိုင်ပါသည် (ပြင်ဆင်ပြီးပါက အောက်ပါ Save Button ကို နှိပ်ပါ)")
    
    editable_df = df_cash[["Date", "Collector", "Customer", "Payment_Type", "Cheque_No", "Amount"]].copy()
    edited_df = st.data_editor(editable_df, num_rows="dynamic", use_container_width=True, key="data_editor")

    if st.button("💾 Save Table Edits"):
        edited_df.to_csv(cash_file, index=False)
        st.success("Changes saved successfully!")
        st.rerun()

else:
    st.info("💡 လက်ရှိတွင် အချက်အလက် စာရင်းများ မရှိသေးပါ၊ Sidebar က Data Entry Form တွင် စာရင်းသစ် စတင်ထည့်သွင်းနိုင်ပါသည်။")
    
