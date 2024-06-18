import pandas as pd
import datetime as dt

entry = pd.read_excel(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\SP QUERY\SP Entry Documents Query-Marco.xlsm")
entry['Entry No.'] = pd.to_numeric(entry['Entry No.'], errors='coerce')
entry = entry[entry['Entry No.'].notna()].astype({'Entry No.': int})
entry = entry.reset_index(drop=True)
entry['ETA'] = entry['ETA'].dt.strftime('%m/%d/%Y')
entry['Modified'] = entry['Modified'].dt.strftime('%m/%d/%Y')
entry.to_excel(r"C:\Users\TechCSG\OneDrive - JH CHB\Shared Documents\Automation\SP QUERY\sp_query_cleaned.xlsx", index=False)

#C:\Users\TechCSG\JH CHB\Data Analysis - Automation\SP QUERY\SP Entry Documents Query-Marco.xlsm
