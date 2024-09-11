import os
import sys
import pandas as pd
from openpyxl import load_workbook

# Define your snippet function to process the data
def process_sheet(df):
    try:
        # Create the mapping dictionary from column 1 and 2
        mapping_dict = pd.Series(df[2].values, index=df[1]).to_dict()

        # Map column 6 using the created dictionary, and compare columns 8 and 9
        df[9] = df[6].map(mapping_dict)
        df[10] = df[9].astype(str) == df[8].astype(str)
    except (KeyError, ValueError) as e:
        # Handle any KeyError, duplicates, or NaN in keys by setting column 10 to False
        df[9] = False
        df[10] = False
    return df

# Function to find the folder containing the Excel files
def get_current_folder():
    if getattr(sys, 'frozen', False):
        # If the script is frozen (turned into an .exe by PyInstaller), use the executable folder
        folder_path = os.path.dirname(sys.executable)
    else:
        # If the script is running normally, use the script's location
        folder_path = os.path.dirname(os.path.abspath(__file__))
    return folder_path

# Function to process all Excel files and sheets in the folder
def process_excel_files(output_file):
    folder_path = get_current_folder()

    # Initialize the result list to store the sheet names and required columns
    result = []

    # Loop through every Excel file in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)
            print(f"Processing file: {file_path}")

            # Load the workbook and process each sheet
            workbook = load_workbook(file_path)
            for sheetname in workbook.sheetnames:
                print(f"Processing sheet: {sheetname}")

                # Load the sheet into a pandas DataFrame
                df = pd.DataFrame(workbook[sheetname].values)

                # Apply the processing logic to the sheet
                df_processed = process_sheet(df)

                # Write back the processed DataFrame to the workbook
                for row in dataframe_to_rows(df_processed, index=False, header=False):
                    workbook[sheetname].append(row)

                # Append sheet name and required columns to the result list
                for index, row in df_processed.iterrows():
                    result.append([sheetname, row[6], row[8], row[9]])

            # Save the workbook with the processed data
            workbook.save(file_path)

    # Save the result list to a CSV file
    result_df = pd.DataFrame(result, columns=['Sheet Name', 'Column 6', 'Column 8', 'Column 9'])
    result_df.to_csv(os.path.join(folder_path, output_file), index=False)
    print(f"Results saved to {output_file}")

if __name__ == "__main__":
    output_file = "output.csv"  # Define your output file name
    process_excel_files(output_file)
