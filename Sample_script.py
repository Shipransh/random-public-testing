# Adjusted transformation function with the correct column name
def transform_toulouse_data_corrected(toulouse_df):
    # Define columns for the transformed dataframe
    transformed_df = pd.DataFrame(columns=[
        'ESN', 'Pos', 'A/C', 'PARTS REPLACED', 'REMOVED P/N', 'REMOVED S/N', 
        'INSTALLED P/N', 'INSTALLED S/N', 'NAME', 'LOCATION', 'DATE'
    ])

    # Set constant values
    name = "name"
    location = "Toulouse"
    date = None  # Empty date

    # Process each row based on structure observed in Toulouse.csv
    for i in range(2, len(toulouse_df)):  # Start from row 2 to skip header rows
        row = toulouse_df.iloc[i]

        # Populate the transformed data row by row
        esn = row['A/C']
        pos = row['A/C-Value']
        part_replaced = row['PARTS REPLACEMENT STATUS']
        removed_pn = row['Unnamed: 3']
        removed_sn = row['Unnamed: 4']
        installed_pn = row['Unnamed: 5']
        installed_sn = None  # Empty as no data available

        # Append transformed row
        transformed_df = transformed_df.append({
            'ESN': esn,
            'Pos': pos,
            'A/C': "A/C-Value",  # Set constant A/C field as per output sample
            'PARTS REPLACED': part_replaced,
            'REMOVED P/N': removed_pn,
            'REMOVED S/N': removed_sn,
            'INSTALLED P/N': installed_pn,
            'INSTALLED S/N': installed_sn,
            'NAME': name,
            'LOCATION': location,
            'DATE': date
        }, ignore_index=True)

    return transformed_df

# Apply the corrected transformation
transformed_df_corrected = transform_toulouse_data_corrected(toulouse_df)

# Save the transformed data to an Excel file
output_file_path_corrected = '/mnt/data/Transformed_Toulouse_Output_Corrected.xlsx'
transformed_df_corrected.to_excel(output_file_path_corrected, index=False)

output_file_path_corrected
