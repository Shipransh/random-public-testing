import os
import sys
import pandas as pd
from openpyxl import load_workbook

def find_excel_files_in_current_folder():
    """Identify all Excel files in the current folder."""
    if getattr(sys, 'frozen', False):
        # If the script is frozen (e.g., turned into an .exe by PyInstaller), use the executable folder
        folder_path = os.path.dirname(sys.executable)
    else:
        # If the script is running normally (not frozen), use the script's location
        folder_path = os.path.dirname(os.path.abspath(__file__))
    
    # List all Excel files in the folder
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    return folder_path, excel_files

def process_excel_files(output_file):
    """Process each Excel file in the folder, applying transformations and collecting data."""
    folder_path, excel_files = find_excel_files_in_current_folder()
    summary_data = []  # To store sheet names and column data

    for excel_file in excel_files:
        file_path = os.path.join(folder_path, excel_file)
        print(f"Processing file: {file_path}")
        
        # Load the Excel file using openpyxl
        excel_data = pd.ExcelFile(file_path, engine='openpyxl')
        
        # Process each sheet in the Excel file
        for sheet_name in excel_data.sheet_names:
            df = pd.read_excel(excel_data, sheet_name=sheet_name, engine='openpyxl')
            print(f"Processing sheet: {sheet_name}")

            # Check if the necessary columns exist in the dataframe
            required_columns = [1, 2, 6, 8]
            if not all(col in df.columns for col in required_columns):
                print(f"Skipping sheet {sheet_name} in {excel_file} - required columns not found.")
                continue

            try:
                # Apply the transformation to the sheet
                mapping_dict = pd.Series(df[2].values, index=df[1]).to_dict()
                df[9] = df[6].map(mapping_dict)
                df[10] = df[9].astype(str) == df[8].astype(str)

                # Save the modified sheet back into the Excel file
                output_file_path = os.path.join(folder_path, f"edited_{excel_file}")
                
                # If the file already exists, load it and append the new sheet
                if os.path.exists(output_file_path):
                    with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Extract relevant columns for the summary
                for _, row in df.iterrows():
                    summary_data.append({
                        'File Name': excel_file,
                        'Sheet Name': sheet_name,
                        'Column 6': row[6],
                        'Column 8': row[8],
                        'Column 9': row[9]
                    })

            except KeyError as e:
                print(f"KeyError {e} encountered in sheet {sheet_name} of file {excel_file}. Skipping sheet.")
                continue

    # Output the collected data into a CSV file
    if summary_data:
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_csv(os.path.join(folder_path, output_file), index=False)
        print(f"Summary saved to {output_file}")
    else:
        print("No data to save to summary file.")

if __name__ == '__main__':
    output_file = "summary_output.csv"
    process_excel_files(output_file)
