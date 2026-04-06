import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


@st.cache_resource
def get_gsheet_client():
    """Initialize gspread client from Streamlit secrets."""
    credentials = Credentials.from_service_account_info(
        st.secrets["connections"]["gsheets"],
        scopes=SCOPES,
    )
    return gspread.authorize(credentials)


st.title("Google Sheet Viewer")

url = st.text_input(
    "Google Sheet URL",
    value="https://docs.google.com/spreadsheets/d/1n2kRtJ6axWrb1YB1VGU3iOXPDusLq9XSbfo2cvhXQnM/edit?usp=sharing",
)

# Lấy danh sách worksheets
try:
    client = get_gsheet_client()
    spreadsheet = client.open_by_url(url.strip())
    worksheets = [ws.title for ws in spreadsheet.worksheets()]
    
    if worksheets:
        # Hiển thị danh sách tất cả worksheets
        st.subheader(f"📊 Danh sach cac bang ({len(worksheets)} worksheet)")
        st.write(", ".join([f"`{ws}`" for ws in worksheets]))
        
        st.divider()
        
        # Dropdown chọn worksheet
        selected_sheet = st.selectbox("Chon bang de xem du lieu", options=worksheets, index=0 if worksheets else None)
        
        # Hiển thị dữ liệu ngay khi chọn
        if selected_sheet:
            st.subheader(f"📋 Du lieu bang: {selected_sheet}")
            try:
                worksheet = spreadsheet.worksheet(selected_sheet)
                rows = worksheet.get_all_values()
                
                if rows:
                    data = pd.DataFrame(rows[1:], columns=rows[0])
                    st.success(f"Da tai bang: {selected_sheet} ({len(rows)-1} dong)")
                    st.dataframe(data, use_container_width=True)
                    
                    # Form data entry
                    st.divider()
                    st.subheader(f"➕ Them du lieu moi vao bang: {selected_sheet}")
                    
                    columns = rows[0]  # Header
                    
                    with st.form(f"form_{selected_sheet}"):
                        form_data = {}
                        for col in columns:
                            form_data[col] = st.text_input(f"{col}", key=f"input_{col}")
                        
                        submitted = st.form_submit_button("Luu du lieu")
                    
                    if submitted:
                        try:
                            # Lấy giá trị từ form
                            new_row = [form_data[col] for col in columns]
                            
                            # Append row vào worksheet
                            worksheet.append_row(new_row, value_input_option="USER_ENTERED")
                            st.success("✅ Da luu du lieu vao Google Sheet!")
                        except Exception as err:
                            st.error(f"Khong the luu du lieu: {err}")
                else:
                    st.warning("Bang trong")
            except Exception as err:
                st.error(f"Khong the tai du lieu: {err}")
                import traceback
                st.text(traceback.format_exc())
    else:
        st.info("Khong co worksheet trong Google Sheet nay.")
        
except Exception as e:
    st.error(f"Khong the tai danh sach worksheet: {e}")
    import traceback
    st.text(traceback.format_exc())
