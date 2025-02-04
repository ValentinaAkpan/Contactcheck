import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re

# Function to load ALL sheets in the Excel file
def load_excel(file):
    all_sheets = pd.read_excel(file, sheet_name=None, dtype=str)  # Load all sheets as dictionary
    combined_data = pd.DataFrame()  # Placeholder for merged data

    # Loop through each sheet and extract company names
    for sheet_name, df in all_sheets.items():
        df = df.fillna('')  # Replace NaN values with empty strings
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)  # Strip spaces

        # Identify the correct column for company names
        possible_names = ["Company", "Company Name", "Business Name", "Firm", "Member Name"]
        found_column = None
        for col in df.columns:
            if col.strip() in possible_names:
                found_column = col.strip()
                break

        if found_column:
            combined_data = pd.concat([combined_data, df[[found_column]]], ignore_index=True)

    # Return a combined unique list of companies
    return set(combined_data.iloc[:, 0].dropna())  # Convert to a set

# Function to extract only left-side company names from Word
def extract_company_names(doc):
    company_names = set()
    lines = [para.text.strip() for para in doc.paragraphs if para.text.strip()]  # Remove empty lines

    i = 0
    while i < len(lines):
        text = lines[i]  # Get first line (assumed company name)

        # Ignore lines with numbers (phone numbers), websites, emails
        if re.search(r'\d', text) or "@" in text or "www." in text or ".com" in text or ".ca" in text:
            i += 1
            continue

        # Add only the first line of the block (left-side company name)
        company_names.add(text)

        # Skip the next 3 lines (address, phone number, contact)
        i += 4

    return company_names

# Streamlit UI
st.title("Membership Directory Checker")
st.write("Upload an **Excel** file and an **existing Word document** to compare directory information.")

# File Uploads
excel_file = st.file_uploader("Upload Updated Excel File", type=["xlsx"])
word_file = st.file_uploader("Upload Existing Word Document", type=["docx"])

if excel_file and word_file:
    st.success("Files uploaded successfully!")

    # Load Excel data
    excel_companies = load_excel(excel_file)

    st.write(f"Total Companies in Excel (from all sheets): {len(excel_companies)}")

    # Load Word file and extract company names correctly
    word_doc = Document(word_file)
    word_companies = extract_company_names(word_doc)

    st.write(f"Total Companies in Word Document: {len(word_companies)}")

    # Identify missing and removed companies
    missing_in_word = excel_companies - word_companies  # Companies in Excel but missing from Word
    extra_in_word = word_companies - excel_companies  # Companies in Word but not in Excel

    st.write("### Missing Companies (Need to be Added)")
    if missing_in_word:
        st.write(f"**Companies in Excel but Missing from Word:** {len(missing_in_word)}")
        st.write(list(missing_in_word))
    else:
        st.write("No new companies missing.")

    st.write("### Companies in Word but Not in Excel (Possible Removals)")
    if extra_in_word:
        st.write(f"**Companies in Word but No Longer in Excel:** {len(extra_in_word)}")
        st.write(list(extra_in_word))
    else:
        st.write("No companies need to be removed.")

    # Prepare CSV report
    report_data = pd.DataFrame({
        "Companies to Add": list(missing_in_word) + [""] * (max(len(missing_in_word), len(extra_in_word)) - len(missing_in_word)),
        "Companies to Remove": list(extra_in_word) + [""] * (max(len(missing_in_word), len(extra_in_word)) - len(extra_in_word))
    })

    output = BytesIO()
    report_data.to_csv(output, index=False)
    output.seek(0)

    st.download_button(
        label="Download Report (CSV)",
        data=output,
        file_name="directory_changes.csv",
        mime="text/csv"
    )
