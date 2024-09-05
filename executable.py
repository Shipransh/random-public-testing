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

                # Separate relevant columns for old and new data
                old_data = df[[2, 3]]  # Old data: columns 2 (Key) and 3 (Value_1)
                new_data = df[[7, 9]]  # New data: columns 7 (Key) and 9 (Value_2)

                # Rename columns to prepare for merging
                old_data.columns = ['Old_Key', 'Value_1']
                new_data.columns = ['New_Key', 'Value_2']

                # Perform the merge based on keys
                merged_data = pd.merge(df, old_data, left_on=7, right_on='Old_Key', how='left')

                # Add the merged data (Value_1 from old data) into column 10
                df[10] = merged_data['Value_1']

                # Compare Value_1 (column 10) and Value_2 (column 9) to generate the match result in column 11
                df[11] = df[9] == df[10]

                # Identify rows where there is no match (False in column 11)
                unmatched_df = df[~df[11]].copy()

                # Store unmatched rows information for the report
                if not unmatched_df.empty:
                    for row in unmatched_df.itertuples():
                        unmatched_rows.append({
                            'Key': row[7],  # New Key
                            'Old Info': row[10],  # Old Value_1 (from lookup)
                            'New Info': row[9],  # New Value_2
                            'Sheet Name': sheet_name,
                            'Row Number': row.Index + 1  # Excel uses 1-based indexing
                        })

                # Write the updated dataframe with columns 10 (Value_1_Lookup) and 11 (Match) back to the Excel file
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
