import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📮 Postal Establishment Data Validator")

# Load time factor file once (server-side, stored permanently)
@st.cache_data
def load_time_factors():
    return pd.read_excel("time_factors.xlsx", engine='openpyxl')

time_factors = load_time_factors()

# User uploads only one file (filled data)
uploaded_data_file = st.file_uploader("📤 Upload Filled Excel Data File (.xls/.xlsx)", type=["xls", "xlsx"])

def load_excel(file):
    file.seek(0)
    try:
        if file.name.endswith('.xlsx'):
            return pd.read_excel(file, engine='openpyxl')
        elif file.name.endswith('.xls'):
            return pd.read_excel(file, engine='xlrd')
        else:
            st.error("Unsupported file format. Please upload a valid Excel file.")
            return None
    except Exception as e:
        st.error(f"Error reading the file {file.name}: {e}")
        return None

if uploaded_data_file:
    data = load_excel(uploaded_data_file)

    if data is not None:
        merged_data = pd.merge(data, time_factors, on='transaction_code', how='left')

        missing_factors = merged_data[merged_data['avg_time_factor'].isna()]
        if not missing_factors.empty:
            st.warning(f"⚠️ Missing time factors for transaction codes: {missing_factors['transaction_code'].unique()}")

        merged_data['computed_time'] = merged_data['item_value'] * merged_data['avg_time_factor']

        st.success("✅ Data processed successfully! Preview below:")
        st.dataframe(merged_data.head(20))

        # Excel download link
        buffer = BytesIO()
        merged_data.to_excel(buffer, index=False, engine='openpyxl')

        st.download_button(
            label="📥 Download Processed Excel File (.xlsx)",
            data=buffer.getvalue(),
            file_name="processed_establishment_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
