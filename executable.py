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

                # Perform XLOOKUP-like operation and comparison
                df['Value_1_Lookup'] = df[7].map(df.set_index(2)[3])  # XLOOKUP equivalent
                df['Match'] = df[9] == df['Value_1_Lookup']  # Compare Value_2 (column 9) with Value_1 (Lookup result)

                # Find unmatched rows
                unmatched_df = df[~df['Match']].copy()  # Rows where Match is False

                # Store information about unmatched rows
                if not unmatched_df.empty:
                    for row in unmatched_df.itertuples():
                        unmatched_rows.append({
                            'Key': row[2],  # Old Key
                            'Old Info': row[3],  # Old Value_1
                            'New Info': row[9],  # New Value_2
                            'Sheet Name': sheet_name,
                            'Row Number': row.Index + 1  # Row index in the file (1-based for Excel)
                        })

                # Write back the modified dataframe with Match column to the same sheet
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
