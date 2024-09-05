import os
import pandas as pd
import openpyxl

# Function to process each Excel file
def process_excel_file(file_path, output_report):
    try:
        # Load the workbook and iterate through all sheets
        workbook = pd.ExcelFile(file_path)
        for sheet_name in workbook.sheet_names:
            df = pd.read_excel(workbook, sheet_name=sheet_name, header=None)
            
            # Check if the DataFrame is large enough to contain the expected columns
            if df.shape[1] < 9:
                print(f"Sheet {sheet_name} in file {file_path} does not have enough columns.")
                continue
            
            # Perform the equivalent of XLOOKUP (merge operation based on key columns)
            df['New_Column'] = df.apply(lambda row: lookup_value(row, df), axis=1)
            
            # Compare the new column with Value_2 (column 9)
            df['Match'] = df[9] == df['New_Column']
            
            # Identify unmatched rows
            unmatched_rows = df[df['Match'] == False]
            
            # Append unmatched rows to the output report
            for idx, row in unmatched_rows.iterrows():
                output_report.append({
                    'Key': row[7],  # Column 7 for key
                    'Old Info': row[2],  # Value_1 from column 3
                    'New Info': row[8],  # Value_2 from column 9
                    'Sheet Name': sheet_name,
                    'Row Number': idx + 1  # Add 1 to match Excel row numbers
                })

            # Save the modified DataFrame with 'Match' column back to the Excel file
            save_sheet_with_match_column(file_path, df, sheet_name)

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

# Function to lookup Value_1 based on Key in Column 7 using Column 2
def lookup_value(row, df):
    try:
        # Get the key from column 7
        key_to_lookup = row[6]
        
        # Find matching row in column 2 and return Value_1 from column 3
        matching_row = df[df[1] == key_to_lookup]
        if not matching_row.empty:
            return matching_row.iloc[0, 2]  # Return Value_1 (column 3)
        return None
    except Exception as e:
        print(f"Error during lookup: {e}")
        return None

# Function to save the modified DataFrame with the 'Match' column back to the file
def save_sheet_with_match_column(file_path, df, sheet_name):
    try:
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
    except Exception as e:
        print(f"Error saving file {file_path}: {e}")

# Main function to process all Excel files in the folder
def process_folder(folder_path):
    output_report = []
    
    # Iterate over all Excel files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
            file_path = os.path.join(folder_path, file_name)
            process_excel_file(file_path, output_report)
    
    # Create a report of unmatched rows
    create_output_report(output_report, folder_path)

# Function to create an output report of unmatched rows
def create_output_report(output_report, folder_path):
    report_df = pd.DataFrame(output_report)
    report_file = os.path.join(folder_path, 'unmatched_report.xlsx')
    report_df.to_excel(report_file, index=False)
    print(f"Unmatched report created at {report_file}")

# Specify the folder path containing the Excel files
folder_path = os.path.dirname(os.path.abspath(__file__))  # Get the directory where the script is running
process_folder(folder_path)
