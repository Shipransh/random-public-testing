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
                # Read each sheet into a dataframe without headers
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

                # Check if the required columns exist (zero-indexed columns)
                if len(df.columns) < 10:
                    print(f"Sheet {sheet_name} does not have the expected number of columns.")
                    continue

                # Separate old and new data for merging
                old_data = df[[1, 2]].copy()  # Old data: columns 1 (Key) and 2 (Value_1)
                old_data.columns = ['Old_Key', 'Value_1']  # Rename for merging

                new_data = df[[6, 8]].copy()  # New data: columns 6 (Key) and 8 (Value_2)
                new_data.columns = ['New_Key', 'Value_2']  # Rename for clarity

                # Merge new data with the old data based on the keys
                merged_data = pd.merge(df, old_data, left_on=6, right_on='Old_Key', how='left')

                # Ensure column 10 is added as the lookup of Value_1 from the old data
                df[9] = merged_data['Value_1']

                # Compare Value_1_Lookup (column 9) and Value_2 (column 8)
                df[10] = df[8] == df[9]

                # Identify rows where the Match is False
                unmatched_df = df[~df[10]].copy()

                # Store unmatched rows for reporting
                if not unmatched_df.empty:
                    for row in unmatched_df.itertuples():
                        unmatched_rows.append({
                            'Key': row[6],  # New Key (column 6)
                            'Old Info': row[9],  # Old Value_1 (from lookup)
                            'New Info': row[8],  # New Value_2
                            'Sheet Name': sheet_name,
                            'Row Number': row.Index + 1  # Row index (1-based for Excel)
                        })

                # Write back the updated dataframe with new columns 9 and 10 to the original Excel file
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

            # Output unmatched rows report
            if unmatched_rows:
                output_unmatched_report(file_path, unmatched_rows)

    except Exception as e:
        print(f"Error processing {file_path}: {e}")

def output_unmatched_report(file_path, unmatched_rows):
    # Create a separate Excel report with unmatched rows
    report_file_path = os.path.splitext(file_path)[0] + "_unmatched_report.xlsx"
    unmatched_df = pd.DataFrame(unmatched_rows)
    
    with pd.ExcelWriter(report_file_path, engine='openpyxl') as writer:
        unmatched_df.to_excel(writer, sheet_name="Unmatched Rows", index=False)
    
    print(f"Unmatched rows report generated: {report_file_path}")

# Main execution
if __name__ == "__main__":
    folder_path = os.path.dirname(os.path.realpath(__file__))  # Same folder as the script
    process_excel_files_in_folder(folder_path)
