import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📮 Postal Establishment Data Validator")

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
        st.error(f"❌ Error reading the file '{file.name}': {e}")
        return None

uploaded_data_file = st.file_uploader("📤 Upload Filled Excel Data (.xls/.xlsx)", type=["xls", "xlsx"])

if uploaded_data_file:
    data = load_excel(uploaded_data_file)
    time_factors = pd.read_excel('time_factors.xlsx', engine='openpyxl')

    if data is not None:
        merged_data = pd.merge(data, time_factors, on='transaction_code', how='left')

        missing_factors = merged_data[merged_data['avg_time_factor'].isna()]
        if not missing_factors.empty:
            st.warning(f"⚠️ Missing time factors for transaction codes: {missing_factors['transaction_code'].unique()}")

        merged_data['Daily Time Taken in minutes'] = merged_data['item_value'] * merged_data['avg_time_factor']

        # Retain only requested columns and reorder explicitly
        final_df = merged_data[[
            "transaction_code",
            "item_description",
            "from_date",
            "transaction_description",
            "Daily Time Taken in minutes"
        ]]

        # Sort by computed column descending
        final_df = final_df.sort_values(by='Daily Time Taken in minutes', ascending=False)

        # Calculate total establishment strength
        establishment_strength = final_df['Daily Time Taken in minutes'].sum() / 240

        st.success(f"✅ Processed successfully! Establishment Strength: **{establishment_strength:.2f}**")

        st.dataframe(final_df.head(20))

        # Prepare Excel file with appended summary
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='ProcessedData')
            workbook = writer.book
            worksheet = writer.sheets['ProcessedData']

            # Write summary clearly at the end of sheet
            summary_row = len(final_df) + 2
            worksheet.cell(row=summary_row, column=4).value = "Grand Total (Establishment Strength)"
            worksheet.cell(row=summary_row, column=5).value = round(establishment_strength, 2)

        st.success("✅ Data processed successfully! Preview below:")
        st.dataframe(final_df.head(20))

        st.info(f"🧮 **Total Establishment Strength:** {establishment_strength:.2f}")

        # Download link
        buffer = BytesIO()
        final_df.to_excel(buffer, index=False, engine='openpyxl')
        st.download_button(
            label="📥 Download Processed Excel File (.xlsx)",
            data=buffer.getvalue(),
            file_name="processed_establishment_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
