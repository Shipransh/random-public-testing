import pandas as pd
import os

# Function to process each sheet in a file
def process_sheet(df, col_lookup_index, col_key_index, col_return_index, col_compare_index):
    # Extract the relevant columns based on the provided indexes
    lookup_column = df.iloc[:, col_lookup_index]  # Equivalent to Column5 (G:G in Excel)
    key_column = df.iloc[:, col_key_index]        # Equivalent to Column2 (B:B in Excel)
    return_column = df.iloc[:, col_return_index]  # Equivalent to Column3 (C:C in Excel)
    compare_column = df.iloc[:, col_compare_index]# Equivalent to Column6 (H:H in Excel)
    
    # Perform the XLOOKUP-like operation using pandas' map function
    match_column = lookup_column.map(pd.Series(return_column.values, index=key_column))
    
    # If 'Match' is NaN, replace with 'NA' (similar to Excel's NA() function)
    match_column = match_column.fillna('NA')
    
    # Compare the result of the lookup with the compare column and store the boolean result
    df['Match'] = match_column == compare_column
    
    # Extract unmatched rows and add additional details for the report
    unmatched_rows = df[df['Match'] == False].copy()
    unmatched_rows['Key'] = unmatched_rows.iloc[:, col_key_index]
    unmatched_rows['Old Info'] = unmatched_rows.iloc[:, col_return_index]
    unmatched_rows['New Info'] = unmatched_rows.iloc[:, col_compare_index]
    unmatched_rows['Row'] = unmatched_rows.index + 1  # Adding 1 to make the row number 1-based

    return unmatched_rows[['Key', 'Old Info', 'New Info', 'Row']]

# Function to process each file
def process_file(file_path, col_lookup_index, col_key_index, col_return_index, col_compare_index):
    try:
        # Load the Excel file without headers
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return pd.DataFrame()  # Return an empty DataFrame if there was an error
    
    # Print out all sheet names to verify all sheets are being recognized
    print(f"Processing file: {file_path}")
    print(f"Found sheets: {xls.sheet_names}")
    
    unmatched_rows_all_sheets = pd.DataFrame()
    
    # Process each sheet in the Excel file
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            unmatched_rows = process_sheet(df, col_lookup_index, col_key_index, col_return_index, col_compare_index)
        
            if not unmatched_rows.empty:
                unmatched_rows['Sheet Name'] = sheet_name
                unmatched_rows_all_sheets = pd.concat([unmatched_rows_all_sheets, unmatched_rows], ignore_index=True)
            
            # Save the modified DataFrame back to the same sheet in the Excel file
            with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        except Exception as e:
            print(f"Error processing sheet {sheet_name} in file {file_path}: {e}")
    
    return unmatched_rows_all_sheets

# Get the directory where the script is running
current_directory = os.path.dirname(os.path.abspath(__file__))

# Define the column indexes (adjust these as needed)
col_lookup_index = 4   # This is the index for Column5 (used in XLOOKUP, equivalent to G:G)
col_key_index = 1      # This is the index for Column2 (used as the lookup array, equivalent to B:B)
col_return_index = 2   # This is the index for Column3 (value to return, equivalent to C:C)
col_compare_index = 5  # This is the index for Column6 (value to compare against, equivalent to compare column)

# Prepare to compile unmatched rows from all files
unmatched_report = pd.DataFrame()

# Loop through all files in the directory
for file_name in os.listdir(current_directory):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(current_directory, file_name)
        unmatched_rows = process_file(file_path, col_lookup_index, col_key_index, col_return_index, col_compare_index)
        
        if not unmatched_rows.empty:
            unmatched_rows['File'] = file_name
            unmatched_report = pd.concat([unmatched_report, unmatched_rows], ignore_index=True)

# Reorder columns for the final report
unmatched_report = unmatched_report[['File', 'Sheet Name', 'Row', 'Key', 'Old Info', 'New Info']]

# Output the report with all unmatched rows
report_file_path = os.path.join(current_directory, 'unmatched_report.xlsx')
unmatched_report.to_excel(report_file_path, index=False)

print(f"All Excel files have been processed. Unmatched report saved to {report_file_path}.")
