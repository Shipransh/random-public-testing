import pandas as pd
import os

# Function to apply the transformation to a dataframe
def apply_transformation(df):
    try:
        # Create mapping dictionary from column 1 and column 2
        mapping_dict = pd.Series(df.iloc[:, 1].values, index=df.iloc[:, 0]).to_dict()

        # Map column 6 using the mapping_dict and assign the result to column 9
        df.iloc[:, 8] = df.iloc[:, 5].map(mapping_dict)

        # Compare column 9 and column 8, result is saved in column 10
        df.iloc[:, 9] = df.iloc[:, 8].astype(str) == df.iloc[:, 7].astype(str)

        return df
    except Exception as e:
        print(f"An error occurred: {e}")
        return df

# Get the current working directory (same folder as the executable)
folder_path = os.path.dirname(os.path.realpath(__file__))

# Loop through all Excel files in the folder
for file in os.listdir(folder_path):
    if file.endswith(".xlsx") or file.endswith(".xls"):
        file_path = os.path.join(folder_path, file)

        # Load the Excel file
        try:
            excel_file = pd.ExcelFile(file_path)
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
                for sheet_name in excel_file.sheet_names:
                    # Load the sheet
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)

                    # Apply the transformation
                    df = apply_transformation(df)

                    # Write the updated dataframe back to the sheet
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            print(f"Failed to process {file}: {e}")

print("Processing complete.")
