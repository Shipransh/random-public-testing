import os
import pandas as pd

# Define the function that processes each DataFrame
def process_dataframe(df):
    # Create the mapping dictionary from columns 1 and 2
    mapping_dict = pd.Series(df.iloc[:, 1].values, index=df.iloc[:, 0]).to_dict()
    
    # Apply the mapping to column 6 and store results in column 9
    df[9] = df[6].map(mapping_dict)
    
    # Compare columns 8 and 9 as strings and store results in column 10
    df[10] = df[9].astype(str) == df[8].astype(str)
    
    # Check if there are any False values in column 10
    if not df[10].all():
        print("False rows found!")
        print(df[df[10] == False])  # Output rows where the comparison is False
    else:
        print("No false rows found.")
    
    return df

# Define the function that processes all sheets in an Excel file
def process_excel_file(file_path):
    # Load the Excel file
    xls = pd.ExcelFile(file_path)
    
    # Process each sheet
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"Processing sheet: {sheet_name} in file: {os.path.basename(file_path)}")
        df = process_dataframe(df)

# Main function to process all Excel files in the same folder as this script
def process_all_excel_files_in_folder():
    # Get the current folder where the script is located
    current_folder = os.path.dirname(os.path.abspath(__file__))
    
    # Iterate through all files in the folder
    for file_name in os.listdir(current_folder):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(current_folder, file_name)
            print(f"Processing file: {file_name}")
            process_excel_file(file_path)
            print(f"Finished processing file: {file_name}")

if __name__ == "__main__":
    process_all_excel_files_in_folder()
