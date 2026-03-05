import os
import re
import pandas as pd
import pdfplumber

def clean_name(name):
    if pd.isna(name):
        return ""
    name = str(name).lower()
    name = re.sub(r"\d+\.\d+[\.\d]*", "", name)
    name = re.sub(r"\(x64\)|\(64-bit\)|\(32-bit\)", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"[^\w\s]", " ", name)
    return " ".join(name.split()).strip()

def extract_pdf_table(pdf_path):
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                data.extend(table)
    if not data:
        return pd.DataFrame(columns=['Software'])
    df = pd.DataFrame(data[1:], columns=data[0])
    df.columns = [c.strip() if c else '' for c in df.columns]
    sw_col = [c for c in df.columns if 'software' in c.lower()]
    if sw_col:
        df.rename(columns={sw_col[0]: 'Software'}, inplace=True)
    elif len(df.columns) > 1:
        df.rename(columns={df.columns[1]: 'Software'}, inplace=True)
    return df[['Software']].dropna()

# Load official list
approved = extract_pdf_table('list of softwares (APPROVED by SAM).pdf')
not_approved = extract_pdf_table('LIST OF  SOFTWARES (NOT APPROVED).pdf')

approved['Status'] = 'Allowed'
not_approved['Status'] = 'Not Allowed'
official = pd.concat([approved, not_approved], ignore_index=True).drop_duplicates('Software')

# Clean names
official['Software'] = official['Software'].str.lower().str.strip()
official_clean = official['Software'].apply(clean_name).tolist()
official_status = official['Status'].tolist()
official_names = official['Software'].tolist()

# Test matching Spotify
test_name = "Spotify"
test_clean = clean_name(test_name)

print(f"Testing: '{test_name}'")
print(f"Cleaned: '{test_clean}'")
print("\nChecking exact/substring matches:")

matches = []
for idx, off_clean in enumerate(official_clean):
    if test_clean in off_clean or off_clean in test_clean:
        matches.append((official_names[idx], official_status[idx], off_clean))

if matches:
    print("FOUND MATCHES:")
    for name, status, cleaned in matches:
        print(f"  - '{name}' (Status: {status}) — cleaned: '{cleaned}'")
else:
    print("  No exact/substring matches found")

# Check if Spotify itself is in the list
spotify_in_official = 'spotify' in [s for s in official['Software'].tolist()]
print(f"\nDirect Spotify entry in official list: {spotify_in_official}")

# Show all approved entries that contain 'spotify'
approved_with_spotify = [s for s in official[official['Status'] == 'Allowed']['Software'].tolist() if 'spotify' in s.lower()]
if approved_with_spotify:
    print(f"Approved entries containing 'spotify': {approved_with_spotify}")
