import os
import pandas as pd
from openpyxl import load_workbook

def process_excel_files_in_folder(folder_path):
    # Loop through all files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(folder_path, file_name)
            process_excel_file(file_path)

def process_excel_file(file_path):
    try:
        # Load the Excel file
        with pd.ExcelFile(file_path) as xls:
            unmatched_rows = []

            for sheet_name in xls.sheet_names:
                # Read each sheet into a dataframe
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

                # Separate the old and new data for merging
                old_data = df[[2, 3]]  # Old data: columns 2 (Key) and 3 (Value_1)
                new_data = df[[7, 9]]  # New data: columns 7 (Key) and 9 (Value_2)

                # Rename columns to make them easier to merge
                old_data.columns = ['Old_Key', 'Value_1']
                new_data.columns = ['New_Key', 'Value_2']

                # Merge the old and new data on their keys
                merged_data = pd.merge(new_data, old_data, left_on='New_Key', right_on='Old_Key', how='left')

                # Add the merged data back to the original dataframe in column 10 for Value_1_Lookup
                df[10] = merged_data['Value_1']

                # Compare Value_1_Lookup (column 10) and Value_2 (column 9) and store the result in column 11
                df[11] = df[9] == df[10]  # True if matched, False if not

                # Find unmatched rows (where Match is False)
                unmatched_df = df[~df[11]].copy()  # Rows where Match is False

                # Store information about unmatched rows for reporting
                if not unmatched_df.empty:
                    for row in unmatched_df.itertuples():
                        unmatched_rows.append({
                            'Key': row[7],  # New Key
                            'Old Info': row[10],  # Old Value_1 (from lookup)
                            'New Info': row[9],  # New Value_2
                            'Sheet Name': sheet_name,
                            'Row Number': row.Index + 1  # Row index in the file (1-based for Excel)
                        })

                # Write back the modified dataframe with new columns (10 and 11) to the same sheet
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

            # Output unmatched rows report
            if unmatched_rows:
                output_unmatched_report(file_path, unmatched_rows)

    except Exception as e:
        print(f"Error processing {file_path}: {e}")

def output_unmatched_report(file_path, unmatched_rows):
    # Output the unmatched rows into a separate report file
    report_file_path = os.path.splitext(file_path)[0] + "_unmatched_report.xlsx"
    unmatched_df = pd.DataFrame(unmatched_rows)
    
    with pd.ExcelWriter(report_file_path, engine='openpyxl') as writer:
        unmatched_df.to_excel(writer, sheet_name="Unmatched Rows", index=False)
    
    print(f"Unmatched rows report generated: {report_file_path}")

# Main execution
if __name__ == "__main__":
    folder_path = os.path.dirname(os.path.realpath(__file__))  # Same folder as the script
    process_excel_files_in_folder(folder_path)
