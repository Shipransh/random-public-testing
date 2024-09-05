import os
import pandas as pd
from openpyxl import load_workbook

# Function to process each Excel file and sheet
def process_excel_file(file_path):
    try:
        # Load the workbook
        wb = load_workbook(filename=file_path, data_only=True)
        unmatched_rows = []
        
        for sheet_name in wb.sheetnames:
            # Read the sheet into a DataFrame
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            # Identify the relevant columns
            index_1_col, key_1_col, value_1_col = 0, 1, 2
            index_2_col, key_2_col, value_2_col = 5, 6, 8

            # Convert all values in the key and value columns to strings
            df[key_1_col] = df[key_1_col].astype(str)
            df[key_2_col] = df[key_2_col].astype(str)
            df[value_1_col] = df[value_1_col].astype(str)
            df[value_2_col] = df[value_2_col].astype(str)

            # Check for duplicates in the Key columns
            duplicate_keys_1 = df[key_1_col].duplicated(keep=False)
            duplicate_keys_2 = df[key_2_col].duplicated(keep=False)

            # Add Match column
            df['Match'] = None  # Initialize Match column with None

            # Identify matching rows where keys and values match and there are no duplicates
            for idx, row in df.iterrows():
                if not duplicate_keys_1[idx] and not duplicate_keys_2[idx]:
                    df.at[idx, 'Match'] = (row[key_1_col] == row[key_2_col]) and (row[value_1_col] == row[value_2_col])
                else:
                    df.at[idx, 'Match'] = None

            # Collect unmatched rows (those where Match is False or None)
            unmatched = df[(df['Match'] == False) | (df['Match'].isna())]
            for idx, row in unmatched.iterrows():
                unmatched_rows.append({
                    'Key': row[key_1_col],
                    'Old Info': row[value_1_col],
                    'New Info': row[value_2_col],
                    'Sheet Name': sheet_name,
                    'Row Number': idx + 2  # Adding 2 because of 0-based index and no header
                })

            # Save the updated sheet back to the file
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
        
        return unmatched_rows

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return []

# Function to generate a report for unmatched rows
def generate_report(unmatched_rows, report_path):
    report_df = pd.DataFrame(unmatched_rows)
    report_df.to_excel(report_path, index=False)
    print(f"Unmatched rows report saved to: {report_path}")

# Main function to process all Excel files in the folder
def process_folder(folder_path):
    unmatched_rows = []
    
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(folder_path, file_name)
            print(f"Processing file: {file_path}")
            unmatched_rows += process_excel_file(file_path)
    
    # Generate the report
    report_path = os.path.join(folder_path, 'unmatched_rows_report.xlsx')
    generate_report(unmatched_rows, report_path)

if __name__ == '__main__':
    # Set the folder path where the Excel files are stored
    folder_path = os.path.dirname(os.path.realpath(__file__))
    
    # Process all files in the folder
    process_folder(folder_path)
