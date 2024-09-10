import os
import sys

def find_empty_files_in_current_folder(output_file):
    if getattr(sys, 'frozen', False):
        # If the script is frozen (e.g., turned into an .exe by PyInstaller), use the executable folder
        folder_path = os.path.dirname(sys.executable)
    else:
        # If the script is running normally (not frozen), use the script's location
        folder_path = os.path.dirname(os.path.abspath(__file__))
    
    empty_files = []
    
    # Check all files in the folder
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        
        # Check if it's a file and if it's empty
        if os.path.isfile(file_path) and os.path.getsize(file_path) == 0:
            empty_files.append(file_name)

    # Write the names of the empty files to the output file
    output_file_path = os.path.join(folder_path, output_file)
    with open(output_file_path, 'w') as f:
        for file_name in empty_files:
            f.write(file_name + '\n')
    
    print(f"Empty files list saved to {output_file_path}")

# Example usage:
output_file = "empty_files.txt"  # Replace with the desired output file name

find_empty_files_in_current_folder(output_file)
