import os
import pandas as pd
import sys

# Get the directory where the script/executable is located
if getattr(sys, 'frozen', False):
    folder_path = os.path.dirname(sys.executable)
else:
    folder_path = os.path.dirname(os.path.abspath(__file__))

# Specify the column indices involved in the comparison
column2_index = 1  # Index of the second column (0-based index)
column3_index = 2  # Index of the third column (0-based index)
column5_index = 4  # Index of the fifth column (0-based index)
column6_index = 5  # Index of the sixth column (0-based index)

# List to store the output details
output_details = []

# Function to process each file
def process_file(file_path):
    try:
        # Load the file into a DataFrame based on the file extension
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, engine='openpyxl')
        else:
            return  # Skip files that are neither CSV nor Excel

        # Ensure the dataframe has enough columns to avoid IndexError
        if max(column2_index, column3_index, column5_index, column6_index) >= len(df.columns):
            print(f"One or more column indices are out of bounds for {file_path}. Skipping this file.")
            return

        # Extract the columns based on indices
        column2 = df.iloc[:, column2_index]
        column3 = df.iloc[:, column3_index]
        column5 = df.iloc[:, column5_index]
        column6 = df.iloc[:, column6_index]

        # Match the values in the specified columns
        df['Match'] = (column2 == column5) & (column3 == column6)

        # Identify rows where the match is False
        false_rows = df[~df['Match']]

        # Add details to the output list if mismatches are found
        if not false_rows.empty:
            for index in false_rows.index:
                output_details.append({'File': os.path.basename(file_path), 'Row': index + 1})

        # Save the updated file with the new 'Match' column
        if file_path.endswith('.csv'):
            output_file = file_path
            df.to_csv(output_file, index=False)
        elif file_path.endswith('.xlsx'):
            output_file = file_path.replace('.xlsx', '_updated.xlsx')
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')

    except Exception as e:
        print(f"Error processing {file_path}: {e}")

# Iterate over all files in the folder
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)
    process_file(file_path)

# Create a summary report of all mismatches found
if output_details:
    output_df = pd.DataFrame(output_details)
    output_report_path = os.path.join(folder_path, 'false_rows_report.csv')
    output_df.to_csv(output_report_path, index=False)
    print(f"Processing complete. Mismatches found and recorded in '{output_report_path}'.")
else:
    print("Processing complete. No mismatches found.")
