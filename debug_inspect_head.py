import pandas as pd

file_path = r"C:\Users\Ahad1503038\Desktop\AIC\CreateAccounting_Create Accounting Report (3).xlsx"
try:
    df = pd.read_excel(file_path, sheet_name='Sheet2', header=None, nrows=10)
    print("Dumping first 10 rows:")
    for i, row in enumerate(df.values):
        row_clean = [str(x) for x in row if pd.notna(x) and str(x) != 'nan']
        print(f"Row {i}: {row_clean}")
except Exception as e:
    print(f"Error: {e}")
