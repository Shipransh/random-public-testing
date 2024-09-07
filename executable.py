import os
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def process_excel_files_in_folder(folder_path):
    # Loop through all files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(folder_path, file_name)
            process_excel_file(file_path)

def process_excel_file(file_path):
    try:
        unmatched_rows = []

        # Use openpyxl to load the workbook (efficient for large files)
        wb = load_workbook(file_path)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            df = pd.DataFrame(ws.values)

            # Ensure the dataframe has the expected number of columns (minimum 9 columns needed)
            if df.shape[1] < 9:
                print(f"Sheet {sheet_name} does not have enough columns. Skipping.")
                continue

            # Separate old and new data for merging
            old_data = df[[1, 2]].copy()  # Old data: columns 1 (Key) and 2 (Value_1)
            old_data.columns = ['Old_Key', 'Value_1']  # Rename for merging

            new_data = df[[6, 8]].copy()  # New data: columns 6 (Key) and 8 (Value_2)
            new_data.columns = ['New_Key', 'Value_2']  # Rename for clarity

            # Merge new data with the old data based on the keys
            merged_data = pd.merge(df, old_data, left_on=6, right_on='Old_Key', how='left')

            # Add the merged old value to the dataframe as column 9
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

            # Write the updated dataframe with new columns 9 and 10 back to the sheet incrementally
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

        # Save the updated workbook
        wb.save(file_path)

        # Output unmatched rows report
        if unmatched_rows:
            output_unmatched_report(file_path, unmatched_rows)

    except Exception as e:
        print(f"Error processing {file_path}: {e}")

def output_unmatched_report(file_path, unmatched_rows):
    # Create a separate Excel report with unmatched rows
    report_file_path = os.path.splitext(file_path)[0] + "_unmatched_report.xlsx"
    unmatched_df = pd.DataFrame(unmatched_rows)

    # Write the report efficiently to a new Excel file
    with pd.ExcelWriter(report_file_path, engine='openpyxl') as writer:
        unmatched_df.to_excel(writer, sheet_name="Unmatched Rows", index=False)

    print(f"Unmatched rows report generated: {report_file_path}")

# Main execution
if __name__ == "__main__":
    folder_path = os.path.dirname(os.path.realpath(__file__))  # Same folder as the script
    process_excel_files_in_folder(folder_path)
