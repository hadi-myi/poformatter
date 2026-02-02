import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.title("PO Excel Formatter")

uploaded_file = st.file_uploader("Upload items.xlsx", type="xlsx")
po_number = st.text_input("PO Number", placeholder="e.g., 304972")
order_date = st.date_input("Order Date")

if st.button("Format & Download") and uploaded_file and po_number and order_date:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    
    # Your logic here (keep cols, rename, add PO/date)
    df_clean = df[["oe_order_item_id", "unit_quantity"]].copy()
    df_clean.columns = ["Item ID", "Qty Ordered"]
    df_clean.insert(0, "PO Number", po_number)
    df_clean.insert(1, "Order Date", order_date.strftime("%Y-%m-%d"))
    
    # Save to bytes for download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_clean.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        from openpyxl.worksheet.table import Table, TableStyleInfo
        tab = Table(displayName="PO_Items", ref=f"A1:D{len(df_clean)+1}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        worksheet.add_table(tab)
    output.seek(0)
    
    st.download_button("Download po_items_ready.xlsx", output, "po_items_ready.xlsx")
    st.success("Ready for SharePoint!")
