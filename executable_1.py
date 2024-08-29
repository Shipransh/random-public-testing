import pandas as pd
import os

# Function to process each sheet in a file
def process_sheet(df):
    # Assuming the columns are fixed as described
    col_key_index_1 = 1    # Column for Key (first appearance)
    col_value_index_1 = 2  # Column for Value_1
    col_key_index_2 = 6    # Column for Key (second appearance)
    col_value_index_2 = 8  # Column for Value_2

    # Group by Key and take the first value for cases where there are duplicates
    key_value_map = df.groupby(df.columns[col_key_index_1])[df.columns[col_value_index_1]].first()

    # Map the Key from the second set to the first set's values
    mapped_values = df.iloc[:, col_key_index_2].map(key_value_map)
    
    # Create a boolean column that checks if Value_1 matches Value_2 based on Key
    df['Match'] = mapped_values == df.iloc[:, col_value_index_2]

    # Extract unmatched rows and add additional details for the report
    unmatched_rows = df[df['Match'] == False].copy()
    unmatched_rows['Key'] = df.iloc[:, col_key_index_1]
    unmatched_rows['Old Info'] = df.iloc[:, col_value_index_1]
    unmatched_rows['New Info'] = df.iloc[:, col_value_index_2]
    unmatched_rows['Row'] = unmatched_rows.index + 1  # Adding 1 to make the row number 1-based

    return unmatched_rows[['Key', 'Old Info', 'New Info', 'Row']]

# Function to process each file
def process_file(file_path):
    try:
        # Load the Excel file with headers
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return pd.DataFrame()  # Return an empty DataFrame if there was an error
    
    unmatched_rows_all_sheets = pd.DataFrame()
    
    # Process each sheet in the Excel file
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Process the sheet
            unmatched_rows = process_sheet(df)
        
            if not unmatched_rows.empty:
                unmatched_rows['Sheet Name'] = sheet_name
                unmatched_rows_all_sheets = pd.concat([unmatched_rows_all_sheets, unmatched_rows], ignore_index=True)
            
            # Save the modified DataFrame back to the same sheet in the Excel file
            with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            print(f"Error processing sheet {sheet_name} in file {file_path}: {e}")
    
    return unmatched_rows_all_sheets

# Get the directory where the script is running
current_directory = os.path.dirname(os.path.abspath(__file__))

# Prepare to compile unmatched rows from all files
unmatched_report = pd.DataFrame()

# Loop through all files in the directory
for file_name in os.listdir(current_directory):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(current_directory, file_name)
        unmatched_rows = process_file(file_path)
        
        if not unmatched_rows.empty:
            unmatched_rows['File'] = file_name
            unmatched_report = pd.concat([unmatched_report, unmatched_rows], ignore_index=True)

# Reorder columns for the final report
unmatched_report = unmatched_report[['File', 'Sheet Name', 'Row', 'Key', 'Old Info', 'New Info']]

# Output the report with all unmatched rows
report_file_path = os.path.join(current_directory, 'unmatched_report.xlsx')
unmatched_report.to_excel(report_file_path, index=False)

print(f"All Excel files have been processed. Unmatched report saved to {report_file_path}.")
