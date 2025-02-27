import streamlit as st
import pandas as pd
import re
import io

# Streamlit UI
st.title("Email Filtering App")
st.write("Upload an Excel file with emails, and we'll filter it for you.")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Define possible column names for office/company information
    possible_columns = [
        "Office Name", "Company Name", "Office Details", "Company Details",
        "Office Names", "Company Names", "Office Detail", "Company Detail"
    ]

    # Find the matching column in the dataset
    office_col = next((col for col in df.columns if col.strip().lower() in 
                       [name.lower() for name in possible_columns]), None)

    # If a matching column is found, filter out rows where it is blank
    if office_col:
        df = df[df[office_col].notna() & df[office_col].astype(str).str.strip().ne("")]

    # Define domain filters
    domain_filters = [
        "gmail.com", "googlemail.com", "yahoo.com", "ymail.com", "rocketmail.com",
        "hotmail.com", "live.com", "outlook.com", "msn.com", "aol.com", "aim.com",
        "verizon.net", "att.net", "sbcglobal.net", "bellsouth.net", "ameritech.net",
        "comcast.net", "charter.net", "cox.net", "protonmail.com", "proton.me",
        "mail.com", "gmx.com", "icloud.com", "me.com", "mac.com", "zoho.com",
        "lycos.com", "fastmail.com", "tutanota.com", "earthlink.net", "citlink.net",
        "ca.rr.com", "roadrunner.com"
    ]

    # Define keyword filters (Ensure exact word match + .edu & .gov domains)
    keyword_filters = [
        "abuse", "admin", "account", "advertise", "support", "webmaster",
        "website", r"\bapp\b", r"\bapps\b", "customer", r"\binfo\b", r"\bsales\b",
        r"\.edu$", r"\.gov$"
    ]

    # Convert keyword filters into a regex pattern for exact match
    keyword_pattern = r"\b(" + "|".join(keyword_filters) + r")\b"

    # Filter based on domain
    mask_domain = df['Email'].str.contains('|'.join(domain_filters), case=False, na=False)

    # Filter based on exact match keywords
    mask_keyword = df['Email'].str.contains(keyword_pattern, case=False, na=False, regex=True)

    # Find duplicate records
    duplicate_df = df[df.duplicated(subset=['Email'], keep=False)]

    # Get filtered datasets
    domain_filtered_df = df[mask_domain]
    keyword_filtered_df = df[mask_keyword]
    other_domains_df = df[~(mask_domain | mask_keyword)]

    # Save results to an Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        domain_filtered_df.to_excel(writer, sheet_name='Domain Filters', index=False)
        keyword_filtered_df.to_excel(writer, sheet_name='Keyword Filters', index=False)
        other_domains_df.to_excel(writer, sheet_name='Other Domains', index=False)
        duplicate_df.to_excel(writer, sheet_name='Duplicates', index=False)
        df.to_excel(writer, sheet_name='Full Dataset', index=False)

    st.success("Filtering complete! Download the file below.")

    # Create download button
    st.download_button(label="Download Filtered Excel File",
                       data=output.getvalue(),
                       file_name="filtered_emails.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

