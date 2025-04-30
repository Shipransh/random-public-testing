import smartsheet

# Config
TOKEN = 'your_api_token_here'
WORKSPACE_ID = 123456789  # Replace with your actual Workspace ID
FILE_PATH = 'path/to/your/file.xlsx'
SHEET_NAME = 'Target Sheet Name'  # Name of the sheet to replace

# Initialize client
smartsheet_client = smartsheet.Smartsheet(TOKEN)

# Step 1: Get all sheets in the workspace
workspace = smartsheet_client.Workspaces.get_workspace(WORKSPACE_ID)
existing_sheet = None

for sheet in workspace.data.sheets:
    if sheet.name == SHEET_NAME:
        existing_sheet = sheet
        break

# Step 2: Delete existing sheet if found
if existing_sheet:
    print(f"Deleting existing sheet: {existing_sheet.name} (ID: {existing_sheet.id})")
    smartsheet_client.Sheets.delete_sheet(existing_sheet.id)

# Step 3: Import Excel file as a new sheet
print(f"Importing new sheet from {FILE_PATH}")
new_sheet = smartsheet_client.Sheets.import_xlsx_sheet(
    FILE_PATH,
    to_workspace_id=WORKSPACE_ID
)

print(f"New sheet uploaded: {new_sheet.data.name} (ID: {new_sheet.data.id})")