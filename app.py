import pandas as pd
import re
from google.colab import files

# Upload Excel file
uploaded = files.upload()
file_name = list(uploaded.keys())[0]  # Get uploaded file name (original)
output_file = file_name.replace(".xlsx", "_filtered.xlsx")  # Modify output file name

# Read the Excel file
df = pd.read_excel(file_name)

# Define possible column names for office/company information
possible_columns = [
    "Office Name", "Company Name", "Office Details", "Company Details",
    "Office Names", "Company Names", "Office Detail", "Company Detail"
]

# Find the matching column in the dataset
office_col = None
for col in df.columns:
    if col.strip().lower() in [name.lower() for name in possible_columns]:
        office_col = col
        break

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
    "lycos.com", "fastmail.com", "tutanota.com", "earthlink.net","citlink.net",
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

# Identify duplicate records
duplicates_df = df[df.duplicated(subset=['Email'], keep=False)]

# Filter based on domain
mask_domain = df['Email'].str.contains('|'.join(domain_filters), case=False, na=False)

# Filter based on **exact match** keywords
mask_keyword = df['Email'].str.contains(keyword_pattern, case=False, na=False, regex=True)

# Get filtered datasets
domain_filtered_df = df[mask_domain]
keyword_filtered_df = df[mask_keyword]

# Get emails **not** found in domain or keyword filters
other_domains_df = df[~(mask_domain | mask_keyword)]

# Save results to a new Excel file
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    domain_filtered_df.to_excel(writer, sheet_name='domain_filters', index=False)
    keyword_filtered_df.to_excel(writer, sheet_name='keyword_filters', index=False)
    other_domains_df.to_excel(writer, sheet_name='other_domains', index=False)
    duplicates_df.to_excel(writer, sheet_name='duplicates', index=False)
    df.to_excel(writer, sheet_name='full_dataset', index=False)

# Download the file
files.download(output_file)
