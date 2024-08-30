import os
import pandas as pd

def process_excel_files(folder_path):
    unmatched_report = []

    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)
            try:
                # Load the Excel file
                excel_file = pd.ExcelFile(file_path)
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

                    # Assume columns as per input structure
                    index_1 = 1
                    key_col_1 = 2
                    value_1_col = 3
                    key_col_2 = 7
                    value_2_col = 8

                    # Initialize the Match column with False
                    df['Match'] = False

                    # Compare the values based on keys
                    for index, row in df.iterrows():
                        key_1 = row[key_col_1]
                        key_2 = row[key_col_2]
                        value_1 = row[value_1_col]
                        value_2 = row[value_2_col]

                        if pd.notna(key_1) and pd.notna(key_2) and key_1 == key_2:
                            df.at[index, 'Match'] = value_1 == value_2
                        else:
                            df.at[index, 'Match'] = False

                        if not df.at[index, 'Match']:
                            unmatched_report.append({
                                'Key': key_1 if pd.notna(key_1) else key_2,
                                'Old Info': value_1,
                                'New Info': value_2,
                                'Sheet Name': sheet_name,
                                'Row Number': index + 1
                            })

                    # Save the updated sheet back to the Excel file
                    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"Failed to process {file_path}: {str(e)}")

    # Create a report DataFrame and save it to an Excel file
    report_df = pd.DataFrame(unmatched_report)
    report_file = os.path.join(folder_path, "unmatched_report.xlsx")
    report_df.to_excel(report_file, index=False)
    print(f"Unmatched report generated: {report_file}")

if __name__ == "__main__":
    folder_path = "./"  # Replace with your folder path
    process_excel_files(folder_path)
