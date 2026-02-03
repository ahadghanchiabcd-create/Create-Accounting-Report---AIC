import pandas as pd
import numpy as np
import os
import re

def process_accounting_report(input_path, output_path):
    print(f"Reading input file: {input_path}")
    
    try:
        # Read the raw sheet (Sheet2), header=None because data starts at variable rows
        df = pd.read_excel(input_path, sheet_name='Sheet2', header=None)
    except ValueError as e:
        print(f"Error: Could not read 'Sheet2'. Available sheets might be different. {e}")
        return
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    processed_rows = []
    error_rows = []
    
    # Context variables
    current_section = "Unknown" 
    current_txn_info = {}
    current_txn_lines = []
    current_errors = []
    
    # Relaxed Regex patterns
    txn_num_pattern = re.compile(r"Transaction Number", re.IGNORECASE)
    # Relaxed to capture "Journal Entries Processed", "Journal Entry Processed" etc.
    section_processed_pattern = re.compile(r"Journal Entr.*Processed", re.IGNORECASE)
    # Relaxed to capture "Journal Entries with Errors", "Journal Entry Errors", "Journal Entries Errored"
    section_error_pattern = re.compile(r"Journal Entr.*Error", re.IGNORECASE)
    
    rows = df.values.tolist()
    
    def flush_transaction():
        nonlocal current_txn_lines, current_errors, current_txn_info, current_section
        
        if not current_txn_lines:
            return
        
        for line in current_txn_lines:
            # Merge txn info (Transaction Number, Date, etc.)
            line.update(current_txn_info)
            
            # Logic for Error Section
            if current_section == 'Error':
                line_num = line.get('Line')
                # Find matching errors for this line
                matching_errors = [e.get('Error Message') for e in current_errors 
                                   if str(e.get('Line', '')).strip() == str(line_num).strip() and pd.notna(e.get('Line'))]
                
                if matching_errors:
                    line['Error'] = " | ".join([str(m) for m in matching_errors if pd.notna(m)])
                else:
                    # Fallback: If no specific line match, attach ALL errors for this transaction
                    # This ensures we don't lose the error message usually at the transaction level
                    all_trans_errors = [e.get('Error Message') for e in current_errors if pd.notna(e.get('Error Message'))]
                    if all_trans_errors:
                         line['Error'] = " | ".join([str(m) for m in all_trans_errors])
                    else:
                        line['Error'] = ""
            
            # Append to appropriate list
            if current_section == 'Processed':
                processed_rows.append(line)
            elif current_section == 'Error':
                error_rows.append(line)
        
        # Reset per-transaction buffers
        current_txn_lines = []
        current_errors = []
        # We reset txn_info because "Transaction Number" usually restarts the block
        current_txn_info = {}

    print("Starting processing...")
    
    i = 0
    while i < len(rows):
        row = rows[i]
        row_str = [str(x) for x in row]
        # Filter out 'nan' string which comes from str(np.nan)
        clean_cells = [x for x in row_str if x != 'nan' and x != 'None']
        row_joined = " ".join(clean_cells)
        
        # 1. Detect Section Headers
        # We start with Error pattern check
        if section_error_pattern.search(row_joined):
            flush_transaction()
            print(f"Found Section: Error (Row {i+1})")
            current_section = 'Error'
            i += 1
            continue
            
        if section_processed_pattern.search(row_joined):
            flush_transaction()
            print(f"Found Section: Processed (Row {i+1})")
            current_section = 'Processed'
            i += 1
            continue
            
        # 2. Detect Transaction Header Block
        if txn_num_pattern.search(row_joined):
            flush_transaction()
            
            # Scan forward to extract context keys (Transaction Number, Event Class, etc.)
            # Stop when we hit the table header "Accounting Class" or proper data
            j = i
            while j < len(rows):
                r_scan = rows[j]
                r_scan_str = [str(x) for x in r_scan]
                
                # Stop if we hit the line header
                if "Accounting Class" in r_scan_str:
                    break
                
                # Stop if we hit a new section or transaction (safety break)
                # But allow the FIRST row (j=i) to match
                if j > i:
                    r_scan_joined = " ".join([str(x) for x in r_scan if str(x) != 'nan'])
                    if txn_num_pattern.search(r_scan_joined):
                        break

                # Extract Key-Value pairs
                # We look for known keys and find the first value to their right
                known_keys = ["Transaction Number", "Event Class", "Event Type", "Ledger", 
                              "Accounting Date", "Transaction Date", "Source"]
                
                for k in known_keys:
                    # Simple check: is key in this row?
                    if k in r_scan_str:
                        # Find the index
                        try:
                            k_idx = r_scan_str.index(k)
                            # Find next non-nan value
                            for val_idx in range(k_idx + 1, len(r_scan)):
                                val = r_scan[val_idx]
                                if pd.notna(val) and str(val).lower() != 'nan' and str(val).strip() != '':
                                    current_txn_info[k] = val
                                    break
                        except ValueError:
                            pass
                j += 1
            
            # Don't advance i completely to j, because j might be the "Accounting Class" header line 
            # which needs to be processed by block 3.
            # But we can advance i to j-1 to skip the context rows we just read?
            # actually checking the while condition: we want to process the row j if it is "Accounting Class"
            i = j 
            # Note: if j broke because of "Accounting Class", i is now at that line.
            # Next iteration will catch it in block 3.
            continue

        # 3. Detect Line Table Header
        if "Accounting Class" in row_str:
            # Parse Header
            line_header_map = {}
            for idx, cell in enumerate(row):
                if pd.notna(cell):
                    line_header_map[idx] = str(cell).strip()
            
            # Consume Data Rows
            i += 1
            while i < len(rows):
                r_data = rows[i]
                r_data_str = [str(x) for x in r_data]
                r_data_joined = " ".join(r_data_str)
                
                # Exit conditions: New Transaction or New Section or Error Header
                if txn_num_pattern.search(r_data_joined) or \
                   section_processed_pattern.search(r_data_joined) or \
                   section_error_pattern.search(r_data_joined):
                    break
                
                # Special Case: "Error Message" header might appear inside an Error section block
                if "Error Message" in r_data_str:
                    break
                
                # Check for "Total for Journal Entry" and skip
                if "Total for Journal Entry" in r_data_joined:
                    i += 1
                    continue

                # Check for validity
                # A valid line usually has a non-empty 'Accounting Class' or 'Line'
                line_obj = {}
                is_valid = False
                for col_idx, col_name in line_header_map.items():
                    if col_idx < len(r_data):
                        val = r_data[col_idx]
                        line_obj[col_name] = val
                        # criteria: Line column exists or Accounting Class exists
                        if (col_name == "Line" and pd.notna(val)) or \
                           (col_name == "Accounting Class" and pd.notna(val)):
                            is_valid = True
                            
                if is_valid:
                    current_txn_lines.append(line_obj)
                
                i += 1
            continue

        # 4. Detect Error Table Header
        if "Error Message" in row_str:
            error_header_map = {}
            for idx, cell in enumerate(row):
                if pd.notna(cell):
                    error_header_map[idx] = str(cell).strip()
            
            i += 1
            while i < len(rows):
                r_err = rows[i]
                r_err_str = [str(x) for x in r_err]
                r_err_joined = " ".join(r_err_str)
                
                if txn_num_pattern.search(r_err_joined) or \
                   section_processed_pattern.search(r_err_joined) or \
                   section_error_pattern.search(r_err_joined):
                    break

                if "Total for Journal Entry" in r_err_joined:
                    i += 1
                    continue
                
                err_obj = {}
                is_valid = False
                for col_idx, col_name in error_header_map.items():
                    if col_idx < len(r_err):
                        val = r_err[col_idx]
                        err_obj[col_name] = val
                        if col_name == "Error Message" and pd.notna(val):
                            is_valid = True
                
                if is_valid:
                    current_errors.append(err_obj)
                    
                i += 1
            continue

        # Advance if nothing matched
        i += 1

    # Final flush
    flush_transaction()
    
    # Save Output
    print(f"Extraction complete. Processed Lines: {len(processed_rows)}, Error Lines: {len(error_rows)}")
    
    df_proc = pd.DataFrame(processed_rows)
    df_err = pd.DataFrame(error_rows)
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if not df_proc.empty:
                df_proc.to_excel(writer, sheet_name='Journal Entries Processed', index=False)
            else:
                pd.DataFrame({'Message': ['No processed entries found']}).to_excel(writer, sheet_name='Journal Entries Processed', index=False)
                
            if not df_err.empty:
                df_err.to_excel(writer, sheet_name='Journal Entries Errored', index=False)
            else:
                pd.DataFrame({'Message': ['No errored entries found']}).to_excel(writer, sheet_name='Journal Entries Errored', index=False)
            
            # Auto-adjust column width
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column_cells in worksheet.columns:
                    length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                    # Add a little padding, max out at 50 to avoid massive columns
                    worksheet.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 60)

        print(f"Successfully saved to: {output_path}")
    except PermissionError:
        print(f"CRITICAL ERROR: Permission denied when writing to '{output_path}'.")
        print("Please close the Excel file if it is open and run the script again.")
    except Exception as e:
        print(f"Error saving file: {e}")

if __name__ == "__main__":
    # Adjust paths as needed or use relative paths
    base_dir = r"C:\Users\Ahad1503038\Desktop\AIC"
    input_file = os.path.join(base_dir, "CreateAccounting_Create Accounting Report (3).xlsx")
    output_file = os.path.join(base_dir, "Output_Accounting_Report_v2.xlsx")
    
    process_accounting_report(input_file, output_file)
