import pandas as pd
import re

file_path = r"C:\Users\Ahad1503038\Desktop\AIC\CreateAccounting_Create Accounting Report (3).xlsx"
try:
    df = pd.read_excel(file_path, sheet_name='Sheet2', header=None)
    for i, row in enumerate(df.values):
        row_str = " ".join([str(x) for x in row if pd.notna(x)])
        if re.search(r"Journal Entries.*Processed", row_str, re.IGNORECASE) or \
           re.search(r"Journal Entries.*Error", row_str, re.IGNORECASE):
            print(f"Row {i}: {row_str}")
except Exception as e:
    print(f"Error: {e}")
