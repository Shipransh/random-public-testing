import os
import pandas as pd
import sys

# Get the directory where the script/executable is located
if getattr(sys, 'frozen', False):
    folder_path = os.path.dirname(sys.executable)
else:
    folder_path = os.path.dirname(os.path.abspath(__file__))

# Specify the column names to compare
column1 = 'Column1_Name'
column2 = 'Column2_Name'

# List to store the output details
output_details = []

# Function to process each file
def process_file(file_path):
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    elif file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)
    else:
        return

    # Create a new column with boolean values
    df['Match'] = df[column1] == df[column2]

    # Identify rows where the match is False
    false_rows = df[~df['Match']]

    # Add details to the output list
    if not false_rows.empty:
        for index, row in false_rows.iterrows():
            output_details.append({'File': os.path.basename(file_path), 'Row': index + 1})

    # Save the updated file with the new Match column
    output_file = file_path if file_path.endswith('.csv') else file_path.replace('.xlsx', '_updated.xlsx')
    df.to_csv(output_file, index=False) if file_path.endswith('.csv') else df.to_excel(output_file, index=False)

# Iterate over all files in the folder
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)
    process_file(file_path)

# Output the list of files and row numbers where False rows are found
output_df = pd.DataFrame(output_details)
output_report_path = os.path.join(folder_path, 'false_rows_report.csv')
output_df.to_csv(output_report_path, index=False)

print(f"Processing complete. Check '{output_report_path}' for details.")
