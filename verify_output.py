import pandas as pd

file_path = r"C:\Users\Ahad1503038\Desktop\AIC\Output_Accounting_Report_v2.xlsx"
try:
    df_err = pd.read_excel(file_path, sheet_name='Journal Entries Errored')
    print("Error Sheet Columns:", df_err.columns.tolist())
    print("First Error Row:", df_err.head(1).to_dict('records'))
    
    df_proc = pd.read_excel(file_path, sheet_name='Journal Entries Processed')
    print("Processed Sheet Rows:", len(df_proc))
except Exception as e:
    print(f"Error: {e}")
