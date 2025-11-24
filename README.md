# Python Excel Automation – Data Cleaning & Report Generator

This project demonstrates how Python can automate data cleaning and Excel report creation.  
It simulates a common workflow found in provider enrollment, credentialing, data operations, and administrative teams where CSV files need to be validated, cleaned, and formatted before reporting.

This project uses:
- Python  
- Pandas  
- ExcelWriter  
- CSV input files  
- Automated data cleaning logic  

---

## Project Overview

The script takes an input CSV file that contains:
- Missing values  
- Duplicates  
- Invalid statuses  
- Inconsistent capitalization  
- Misformatted dates  

The Python script:
1. Loads and inspects the data  
2. Removes duplicates  
3. Fixes missing fields where possible  
4. Standardizes text formats (e.g., Active vs active vs ACTIVE)  
5. Converts dates into a consistent format  
6. Writes a clean and ready-to-use Excel report

**My Notes:**  
I built this to practice Python automation for data cleaning, similar to workflows in credentialing, admin operations, and insurance data processing.

---

## Technologies Used

| Category | Tool |
|----------|-------|
| Language | Python |
| Libraries | pandas, openpyxl |
| Formats | CSV, XLSX |
| Platform | Windows/Mac/Linux |

---

## Folder Structure
python-excel-automation-data-cleaning/
│
├─ README.md
├─ scripts/
│ ├─ clean_data.py
│
├─ input_data/
│ ├─ provider_data_raw.csv
│
└─ output/
├─ provider_data_clean.xlsx


---

## Sample Input Data

File: `input_data/provider_data_raw.csv`

```csv
ProviderID,FirstName,LastName,Status,StartDate
1,Sarah,Coleman,active,2022/01/10
2,James,Lee,ACTIVE,2021-11-05
3,Mia,Davis,Inactive,2020.06.18
3,Mia,Davis,Inactive,2020.06.18
4,,Nguyen,Active,2022-09-01


import pandas as pd

# Load CSV
df = pd.read_csv("./input_data/provider_data_raw.csv")

# Standardize column names
df.columns = df.columns.str.lower()

# Remove duplicates
df = df.drop_duplicates()

# Fix capitalization
df['status'] = df['status'].str.strip().str.upper()

# Standardize dates
df['startdate'] = pd.to_datetime(df['startdate'], errors='coerce')

# Fill missing first names with "Unknown"
df['firstname'] = df['firstname'].fillna("Unknown")

# Sort the data
df = df.sort_values(by='providerid')

# Export to Excel
output_path = "./output/provider_data_clean.xlsx"
df.to_excel(output_path, index=False)

print("Data cleaning complete. File saved to:", output_path)

*HOW TO RUN*
python ./scripts/clean_data.py

Results
-Duplicates removed
-Formatting standardized
-Missing values handled
-Dates fixed
-Final Excel report created

**Personal Reflection**
This project taught me how to automate repetitive data cleaning tasks, which is extremely helpful in credentialing, data entry, and provider enrollment workflows. I learned how to use Pandas to transform messy datasets into reliable, clean reports.

Author
Telma Anika
Email: tellyannika@gmail.com
