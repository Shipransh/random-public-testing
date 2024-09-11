import os
import sys
import pandas as pd

def find_excel_files_in_current_folder():
    """Identify all Excel files in the current folder."""
    if getattr(sys, 'frozen', False):
        # If the script is frozen (e.g., turned into an .exe by PyInstaller), use the executable folder
        folder_path = os.path.dirname(sys.executable)
    else:
        # If the script is running normally (not frozen), use the script's location
        folder_path = os.path.dirname(os.path.abspath(__file__))
    
    # List all Excel files in the folder
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]
    return folder_path, excel_files

def process_excel_files(output_file):
    """Process each Excel file in the folder, applying transformations and collecting data."""
    folder_path, excel_files = find_excel_files_in_current_folder()
    summary_data = []  # To store sheet names and column data

    for excel_file in excel_files:
        file_path = os.path.join(folder_path, excel_file)
        print(f"Processing file: {file_path}")
        
        # Load the Excel file
        excel_data = pd.ExcelFile(file_path)
        
        # Process each sheet in the Excel file
        for sheet_name in excel_data.sheet_names:
            df = pd.read_excel(excel_data, sheet_name=sheet_name)
            print(f"Processing sheet: {sheet_name}")

            # Apply the transformation to the sheet
            mapping_dict = pd.Series(df[2].values, index=df[1]).to_dict()
            df[9] = df[6].map(mapping_dict)
            df[10] = df[9].astype(str) == df[8].astype(str)

            # Save the modified sheet back into the Excel file
            output_file_path = os.path.join(folder_path, f"edited_{excel_file}")
            with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a') as writer:
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

    # Output the collected data into a CSV file
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_csv(os.path.join(folder_path, output_file), index=False)
    print(f"Summary saved to {output_file}")

if __name__ == '__main__':
    output_file = "summary_output.csv"
    process_excel_files(output_file)
