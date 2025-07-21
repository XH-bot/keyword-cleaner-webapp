import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook

def clean_keyword(keyword):
    if pd.isna(keyword):
        return ""
    return keyword.strip().lower().replace("--", "-").replace("  ", " ")

def process_file(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_keywords = []
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)
            for col in df.columns:
                if df[col].dtype == object:
                    all_keywords += df[col].dropna().astype(str).tolist()
        cleaned = list(set([clean_keyword(k) for k in all_keywords if clean_keyword(k)]))
        return sorted(cleaned)
    except Exception as e:
        st.error(f"Error processing file: {e}")
        return []

def to_excel(keywords):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df = pd.DataFrame({'Cleaned Keywords': keywords})
    df.to_excel(writer, index=False, sheet_name='Cleaned')
    writer.close()
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Keyword Cleaner", layout="centered")
    st.title("🔍 Keyword Cleaner Tool")
    st.write("Upload your Excel keyword files to clean and deduplicate them.")

    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

    if uploaded_file:
        with st.spinner("Processing..."):
            cleaned_keywords = process_file(uploaded_file)
            st.success(f"Cleaned {len(cleaned_keywords)} unique keywords.")
            st.write(cleaned_keywords)  # preview

            output = to_excel(cleaned_keywords)
            st.download_button(
                label="📥 Download Cleaned Keywords",
                data=output,
                file_name="cleaned_keywords.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
