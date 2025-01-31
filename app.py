import pandas as pd
from docx import Document
import streamlit as st

import pandas as pd
import requests
from io import BytesIO

# Load Excel file from GitHub
def load_data():
    url = "https://raw.githubusercontent.com/YOUR_GITHUB_USERNAME/YOUR_REPO/main/aircraft_parts.xlsx"
    response = requests.get(url)
    if response.status_code == 200:
        return pd.ExcelFile(BytesIO(response.content), engine="openpyxl")
    else:
        raise ValueError("Failed to load Excel file from GitHub.")


# Read relevant sheets
def filter_data(excel_data, selected_aircraft):
    parts_df = excel_data.parse('Parts')  # Sheet containing parts
    mro_df = excel_data.parse('MRO')      # Sheet containing MRO capabilities
    
    # Filter based on selected aircraft models
    filtered_parts = parts_df[parts_df['Aircraft Model'].isin(selected_aircraft)]
    filtered_mro = mro_df[mro_df['Aircraft Model'].isin(selected_aircraft)]
    
    return filtered_parts, filtered_mro

from docx import Document

def generate_document(selected_aircraft, parts, mro):
    # Load the existing template
    doc = Document("template.docx")
    
    # Replace placeholders with actual data
    for para in doc.paragraphs:
        if "{{aircraft_models}}" in para.text:
            para.text = para.text.replace("{{aircraft_models}}", ", ".join(selected_aircraft))
        elif "{{parts_list}}" in para.text:
            parts_text = "\n".join([f"- {row['Part Number']}: {row['Description']}" for _, row in parts.iterrows()])
            para.text = para.text.replace("{{parts_list}}", parts_text)
        elif "{{mro_list}}" in para.text:
            mro_text = "\n".join([f"- {row['Capability']} (Location: {row['Facility']})" for _, row in mro.iterrows()])
            para.text = para.text.replace("{{mro_list}}", mro_text)

    # Save the output document
    file_path = "Sales_Collateral.docx"
    doc.save(file_path)
    return file_path


# Streamlit Web App
def main():
    st.title("Aircraft Sales Collateral Generator")
    
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    
    if uploaded_file:
        excel_data = load_data(uploaded_file)
        aircraft_models = list(excel_data.parse('Parts')['Aircraft Model'].unique())
        
        selected_aircraft = st.multiselect("Select Aircraft Models:", aircraft_models)
        
        if st.button("Generate Document"):
            if selected_aircraft:
                parts, mro = filter_data(excel_data, selected_aircraft)
                doc_path = generate_document(selected_aircraft, parts, mro)
                st.success("Document generated successfully!")
                with open(doc_path, "rb") as file:
                    st.download_button("Download Sales Collateral", file, file_name=doc_path)
            else:
                st.warning("Please select at least one aircraft model.")

if __name__ == "__main__":
    main()
