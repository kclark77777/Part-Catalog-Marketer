import pandas as pd
import requests
from io import BytesIO
import streamlit as st
import openpyxl  # Ensure openpyxl is available for reading .xlsx files
from docx import Document

# Load Excel file from GitHub
def load_data():
    url = "https://raw.githubusercontent.com/kclark77777/Part-Catalog-Marketer/main/aircraft-parts.xlsx"
    response = requests.get(url, stream=True)
    if response.status_code == 200:
        content_type = response.headers.get("Content-Type", "")
        if "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" not in content_type:
            raise ValueError("Downloaded file is not an Excel file. Check the GitHub URL.")
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

# Generate a Word document
def generate_document(selected_aircraft, parts, mro):
    doc = Document("template.docx")  # Use a pre-existing template
    
    for para in doc.paragraphs:
        if "{{aircraft_models}}" in para.text:
            para.text = para.text.replace("{{aircraft_models}}", ", ".join(selected_aircraft))
        elif "{{parts_list}}" in para.text:
            parts_text = "\n".join([f"- {row['Part Number']}: {row['Description']}" for _, row in parts.iterrows()])
            para.text = para.text.replace("{{parts_list}}", parts_text)
        elif "{{mro_list}}" in para.text:
            mro_text = "\n".join([f"- {row['Capability']} (Location: {row['Facility']})" for _, row in mro.iterrows()])
            para.text = para.text.replace("{{mro_list}}", mro_text)
    
    file_path = "Sales_Collateral.docx"
    doc.save(file_path)
    return file_path

# Streamlit Web App
def main():
    st.title("Aircraft Sales Collateral Generator")
    
    try:
        excel_data = load_data()
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
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")

if __name__ == "__main__":
    main()
