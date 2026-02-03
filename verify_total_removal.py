import pandas as pd

file_path = r"C:\Users\Ahad1503038\Desktop\AIC\Output_Accounting_Report_v2.xlsx"
try:
    df_err = pd.read_excel(file_path, sheet_name='Journal Entries Errored')
    # Check if any row has "Total for Journal Entry" in any column
    total_found = False
    for i, row in enumerate(df_err.values):
        row_str = " ".join([str(x) for x in row])
        if "Total for Journal Entry" in row_str:
            print(f"FAILED: Found 'Total for Journal Entry' in Error sheet row {i}: {row_str}")
            total_found = True
            break
            
    if not total_found:
        print("PASSED: No 'Total for Journal Entry' found in Error sheet.")
        
    df_proc = pd.read_excel(file_path, sheet_name='Journal Entries Processed')
    total_found_proc = False
    for i, row in enumerate(df_proc.values):
        row_str = " ".join([str(x) for x in row])
        if "Total for Journal Entry" in row_str:
            print(f"FAILED: Found 'Total for Journal Entry' in Processed sheet row {i}: {row_str}")
            total_found_proc = True
            break
            
    if not total_found_proc:
        print("PASSED: No 'Total for Journal Entry' found in Processed sheet.")
        
except Exception as e:
    print(f"Error: {e}")
