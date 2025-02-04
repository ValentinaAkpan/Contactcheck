import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from thefuzz import fuzz

# List of provinces and general words to ignore
EXCLUDE_WORDS = {
    "ontario", "new brunswick", "manitoba", "nova scotia", "northwest territories",
    "saskatchewan", "quebec", "new foundland", "british columbia", "alberta", "branches"
}

# Function to normalize company names (handles minor variations)
def normalize_name(name):
    if not isinstance(name, str):
        return ""
    name = name.strip().lower()  # Remove spaces and convert to lowercase
    name = name.replace("&", "and")  # Replace '&' with 'and'
    name = re.sub(r'[^\w\s]', '', name)  # Remove punctuation
    return name

# Function to load ALL sheets in the Excel file
def load_excel(file):
    all_sheets = pd.read_excel(file, sheet_name=None, dtype=str)  # Load all sheets as dictionary
    combined_data = pd.DataFrame()  # Placeholder for merged data

    # Loop through each sheet and extract company names
    for sheet_name, df in all_sheets.items():
        df = df.fillna('')  # Replace NaN values with empty strings
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)  # Trim spaces

        # Identify the correct column for company names
        possible_names = ["Company", "Company Name", "Business Name", "Firm", "Member Name"]
        found_column = None
        for col in df.columns:
            if col.strip() in possible_names:
                found_column = col.strip()
                break

        if found_column:
            combined_data = pd.concat([combined_data, df[[found_column]]], ignore_index=True)

    # Return a set of normalized company names
    return {normalize_name(name) for name in combined_data.iloc[:, 0].dropna()}

# Function to extract only company names from the Word document (left side, first line only)
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

        # Ignore provinces and general words
        normalized_text = text.lower().strip()
        if normalized_text in EXCLUDE_WORDS:
            i += 1
            continue

        # Ignore short words (likely names or irrelevant text)
        if len(normalized_text.split()) <= 2:  # Likely a personal name
            i += 1
            continue

        # Normalize and add only the first line of each block
        company_names.add(text)

        # Skip the next 3 lines (address, phone number, contact)
        i += 4

    return company_names

# Function to perform fuzzy matching (handles slight name variations)
def find_closest_match(company, company_list):
    best_match = None
    highest_score = 0
    for comp in company_list:
        score = fuzz.ratio(company, comp)  # Compute similarity score
        if score > highest_score:
            highest_score = score
            best_match = comp
    return best_match if highest_score >= 85 else None  # Match if score is 85% or higher

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

    # Load Word file and extract company names
    word_doc = Document(word_file)
    word_companies = extract_company_names(word_doc)

    st.write(f"Total Companies in Word Document: {len(word_companies)}")

    # Identify missing companies using exact and fuzzy matching
    missing_in_word = set()
    for company in excel_companies:
        if company not in word_companies:
            match = find_closest_match(company, word_companies)
            if match is None:
                missing_in_word.add(company)  # Only add if no good match is found

    # Identify extra companies in Word that are not in Excel
    extra_in_word = set()
    for company in word_companies:
        if company not in excel_companies:
            match = find_closest_match(company, excel_companies)
            if match is None:
                extra_in_word.add(company)  # Only add if no good match is found

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
