## Task

Create a Python script to process Excel files in a folder, which may contain over a million records across multiple sheets. The script should:

### 1. Identify Corresponding Values

- Compare `Value_1` (column 3) with `Value_2` (column 8) based on the corresponding `Key` (column 2 and column 7).
- Add a new column labeled `Match` in the Excel sheets, which stores boolean values indicating whether `Value_1` and `Value_2` match based on the `Key`.

### 2. Report Unmatched Rows

- Create an output report that lists all unmatched rows across all sheets in the format:
  - **Key** (from column 2 or 7)
  - **Old Info** (`Value_1` from column 3)
  - **New Info** (`Value_2` from column 8)
  - **Sheet Name**
  - **Row Number**

### 3. Editable & Scalable

- The script should use index-based references for the columns since the data has no headers. Make it easy to change these index references if the data's shape changes in the future.

### 4. Handle All Sheets

- Ensure the script processes all sheets in the Excel files and adds the `Match` column to each sheet with a proper header.

### 5. Error Handling

- Include robust error handling to manage issues such as:
  - Non-unique index values.
  - Errors when reading or writing to the Excel files.
  - Ensure that no duplicate rows are grouped or ignored; any unexpected behavior should result in a `False` value in the `Match` column.

### 6. Excel Formula Equivalent

- The script should perform a function equivalent to the Excel formula `XLOOKUP(G:G, B:B, C:C, NA(), 0, 1)` but adapted for Python, ensuring edits are applied directly to the files.

### 7. Assumptions

- The data does not have headers, so the script should use indexes.
- All Excel files are in the same folder as the executable script.

### 8. Handle File Processing Issues

- Correct issues related to zipfile errors (`BadZipFile: File is not a zip file`), file header issues (`Bad magic number`), and others that may arise during processing.

### 9. No Grouping of Duplicates

- Ensure that the script does not group duplicates; instead, it should add `False` to the `Match` column for any duplicate or unexpected behavior.

## Input Data Structure

- Columns: `index_1` (column 1), `Key` (columns 2 and 7), `Value_1` (column 3), `Unused_1`, `Unused_2`, `Index_2` (column 6), `Unused_3`, `Value_2` (column 8), `Unused_4`

## Output Requirements

- Add a new column labeled `Match`.
- Create a report listing unmatched rows with the specified format.
