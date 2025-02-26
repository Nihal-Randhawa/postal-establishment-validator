import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📮 Postal Establishment Data Validator")

# Cached loading of reference data for better performance
@st.cache_data
def load_time_factors():
    return pd.read_excel('time_factors.xlsx', engine='openpyxl')

def load_excel(file):
    file.seek(0)
    try:
        if file.name.endswith('.xlsx'):
            return pd.read_excel(file, engine='openpyxl')
        elif file.name.endswith('.xls'):
            return pd.read_excel(file, engine='xlrd')
    except Exception as e:
        st.error(f"❌ Error reading '{file.name}': {e}")
        return None

uploaded_data_file = st.file_uploader("📤 Upload Filled Excel File (.xls/.xlsx)", type=["xls", "xlsx"])

if uploaded_data_file:
    data = load_excel(uploaded_data_file)
    time_factors = pd.read_excel('time_factors.xlsx', engine='openpyxl')

    if data is not None:
        merged_data = pd.merge(data, time_factors, on='transaction_code', how='left')

        missing_factors = merged_data[merged_data['avg_time_factor'].isna()]
        if not missing_factors.empty:
            st.warning(f"⚠️ Missing factors for transaction codes: {missing_factors['transaction_code'].unique()}")

        # Calculate daily time
        merged_data['Daily Time Taken in minutes'] = merged_data['item_value'] * merged_data['avg_time_factor']/1500

        # Retain and reorder specified columns clearly
        final_df = merged_data[[
            "transaction_code",
            "item_description",
            "from_date",
            "transaction_description",
            "item_value",
            "avg_time_factor",
            "Daily Time Taken in minutes"
        ]]

        # Sort descending by Daily Time Taken
        final_df = final_df.sort_values(by='Daily Time Taken in minutes', ascending=False)

        # Establishment strength calculation
        establishment_strength = final_df['Daily Time Taken in minutes'].sum() / 240

        st.success("✅ Data processed successfully! Preview below:")
        st.dataframe(final_df.head(20))

        st.info(f"🧮 **Total Hours of Establishment Time Calculated:** {establishment_strength:.2f}")

        # Append total establishment strength as the last row clearly
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='ProcessedData')

            workbook = writer.book
            worksheet = writer.sheets['ProcessedData']

            # Append Grand total below clearly
            summary_row = len(final_df) + 2
            worksheet.cell(row=summary_row, column=4).value = "Grand Total (Establishment Strength)"
            worksheet.cell(row=summary_row, column=6).value = round(establishment_strength, 2)

        st.success(f"✅ **Total Establishment Strength:** {establishment_strength:.2f}")

        st.dataframe(final_df.head(20))

        # Download processed file
        st.download_button(
            label="📥 Download Processed Excel File (.xlsx)",
            data=buffer.getvalue(),
            file_name="processed_establishment_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
