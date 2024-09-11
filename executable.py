import os
import pandas as pd

# Define the function that processes each DataFrame and collects False rows
def process_dataframe(df):
    # Create the mapping dictionary from columns 1 and 2
    mapping_dict = pd.Series(df.iloc[:, 1].values, index=df.iloc[:, 0]).to_dict()
    
    # Apply the mapping to column 6 and store results in column 9
    df[9] = df[6].map(mapping_dict)
    
    # Compare columns 8 and 9 as strings and store results in column 10
    df[10] = df[9].astype(str) == df[8].astype(str)
    
    # Return the rows where df[10] is False
    false_rows = df[df[10] == False]
    return false_rows

# Define the function that processes all sheets in an Excel file and collects all False rows
def process_excel_file(file_path, false_rows_collector):
    # Load the Excel file
    xls = pd.ExcelFile(file_path)
    
    # Process each sheet
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        false_rows = process_dataframe(df)
        
        # If there are any false rows, add them to the collector with the sheet name
        if not false_rows.empty:
            false_rows['Sheet'] = sheet_name  # Add a column with the sheet name
            false_rows_collector.append(false_rows)

# Main function to process all Excel files in the same folder as this script
def process_all_excel_files_in_folder():
    # Get the current folder where the script is located
    current_folder = os.path.dirname(os.path.abspath(__file__))
    
    # Collector to gather all false rows from all sheets
    false_rows_collector = []
    
    # Iterate through all files in the folder
    for file_name in os.listdir(current_folder):
        if file_name.endswith('.xlsx') and '_processed' in file_name:
            file_path = os.path.join(current_folder, file_name)
            print(f"Processing file: {file_name}")
            process_excel_file(file_path, false_rows_collector)
            print(f"Finished processing file: {file_name}")
    
    # If there are any false rows, save them to a new Excel file
    if false_rows_collector:
        combined_false_rows = pd.concat(false_rows_collector, ignore_index=True)
        output_file_path = os.path.join(current_folder, "false_rows_collected.xlsx")
        combined_false_rows.to_excel(output_file_path, index=False)
        print(f"False rows saved to: {output_file_path}")
    else:
        print("No false rows found.")

if __name__ == "__main__":
    process_all_excel_files_in_folder()
