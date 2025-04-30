import pandas as pd
import smartsheet

# Replace with your actual API token and sheet ID
SMARTSHEET_API_TOKEN = 'YOUR_API_TOKEN'
EXISTING_SHEET_ID = 1234567890123456  # Replace with your actual sheet ID

# Sample DataFrame with matching column names
df = pd.DataFrame({
    'column1': ['Alice', 'Bob', 'Charlie'],
    'column2': [30, 25, 35],
    'column3': ['HR', 'Engineering', 'Marketing'],
    'column4': ['NY', 'CA', 'TX']
})

# Create a Smartsheet client
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_API_TOKEN)
smartsheet_client.errors_as_exceptions(True)

# Load the existing sheet
sheet = smartsheet_client.Sheets.get_sheet(EXISTING_SHEET_ID).result

# Map Smartsheet column titles to their IDs
columns = {col.title: col.id for col in sheet.columns}

# Verify columns exist
missing_cols = set(df.columns) - set(columns)
if missing_cols:
    raise ValueError(f"These DataFrame columns are not in the Smartsheet: {missing_cols}")

# Convert DataFrame rows to Smartsheet rows
rows = []
for _, row_data in df.iterrows():
    cells = [
        smartsheet.models.Cell({
            'column_id': columns[col_name],
            'value': row_data[col_name]
        }) for col_name in df.columns
    ]
    row = smartsheet.models.Row({
        'to_bottom': True,
        'cells': cells
    })
    rows.append(row)

# Add rows to the sheet
response = smartsheet_client.Sheets.add_rows(EXISTING_SHEET_ID, rows)
print(f"Successfully added {len(response.result)} rows to sheet ID {EXISTING_SHEET_ID}")