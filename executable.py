def add_headers_to_sheet(df, headers):
    df.columns = headers
    return df

def force_overwrite_excel_sheet(df, file_path, sheet_name):
    try:
        # Try to load the existing workbook
        book = load_workbook(file_path)
        
        # If the sheet exists, remove it
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        
        # Save the workbook after removing the sheet
        book.save(file_path)
        book.close()
    except FileNotFoundError:
        # If the file does not exist, it will be created
        pass

    # Write the DataFrame to the Excel file, creating the sheet
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


# Folder path containing the Excel files
folder_path = r""
# Define the headers you want to add

# Loop through all Excel files in the folder
for file in os.listdir(folder_path):
    try:
        if file.endswith(".xlsx") or file.endswith(".xls"):
            file_path = os.path.join(folder_path, file)
            try:
                excel_file = pd.ExcelFile(file_path)
            except Exception as e:
                print(f"Failed to load Excel file {file}: {e}. Skipping file.")
                continue
            for sheet_name in excel_file.sheet_names:
                try:
                    print(sheet_name)
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                    df = add_headers_to_sheet(df, np.arange(df.shape[1]))
                    mapping_dict = pd.Series(df[2].values, index=df[1]).to_dict()
                    df[9] = df[6].map(mapping_dict)
                    df['Header 10'] = df[9].astype(str) == df[8].astype(str)
                    force_overwrite_excel_sheet(df, file_path=file_path, sheet_name=sheet_name)
                    print(f"Sheet Processed: {sheet_name}")
                except Exception as e:
                    print(f"Failed to process sheet {sheet_name} in file {file}: {e}. Skipping sheet.")
                    continue
    except Exception as e:
        print(f"Failed to process file {file}: {e}. Skipping file.")
        continue
