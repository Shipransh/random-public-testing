import os
import pandas as pd
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

# Define the folder containing Excel files
folder_path = './excel_files/'  # Change to your folder path

# Define the column indices (1-based indexing)
KEY_COL_1 = 1  # Key in column 2
VALUE_1_COL = 2  # Value_1 in column 3
KEY_COL_2 = 6  # Key in column 7
VALUE_2_COL = 8  # Value_2 in column 9

# Output report file for unmatched rows
output_report_file = 'unmatched_rows_report.csv'

# Initialize a DataFrame to collect all unmatched rows
unmatched_rows = pd.DataFrame(columns=['Key', 'Old Info', 'New Info', 'Sheet Name', 'Row Number'])

def process_excel_file(file_path):
    try:
        # Load the Excel file
        excel_file = pd.ExcelFile(file_path)
        print(f'Processing file: {file_path}')
        
        # Loop through all sheets
        for sheet_name in excel_file.sheet_names:
            print(f'Processing sheet: {sheet_name}')
            # Read the sheet into a DataFrame (no headers)
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            # Handle cases where the sheet is empty or has no relevant data
            if df.empty or df.shape[1] < 9:
                print(f'Sheet {sheet_name} has insufficient data or is empty.')
                continue

            # Perform the XLOOKUP equivalent - lookup Value_1 based on Key_1 and Key_2
            df['Lookup_Value'] = df.apply(lambda row: row[VALUE_1_COL] if row[KEY_COL_1] == row[KEY_COL_2] else None, axis=1)

            # Compare Value_2 (column 9) with Lookup_Value (column 10)
            df['Match'] = df.apply(lambda row: row[VALUE_2_COL] == row['Lookup_Value'], axis=1)

            # Save the modified DataFrame back to the same sheet with the new 'Match' column
            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            
            # Collect unmatched rows for reporting
            unmatched = df[df['Match'] == False].copy()
            unmatched['Sheet Name'] = sheet_name
            unmatched['Row Number'] = unmatched.index + 1
            unmatched = unmatched[[KEY_COL_2, VALUE_1_COL, VALUE_2_COL, 'Sheet Name', 'Row Number']]
            unmatched.columns = ['Key', 'Old Info', 'New Info', 'Sheet Name', 'Row Number']
            unmatched_rows = pd.concat([unmatched_rows, unmatched], ignore_index=True)

    except InvalidFileException as e:
        print(f'Error processing file {file_path}: {e}')
    except Exception as e:
        print(f'An unexpected error occurred: {e}')

def main():
    # Loop through all Excel files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            process_excel_file(file_path)
    
    # Save the unmatched rows report
    unmatched_rows.to_csv(output_report_file, index=False)
    print(f'Unmatched rows report saved to {output_report_file}')

if __name__ == '__main__':
    main()
