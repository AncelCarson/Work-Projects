# Author: ChatGPT
# Orginization: Napps Technology Comporation
# Creation Date: 15/10/2024
# Update Date: 15/10/2024
# File Copy.py

import shutil
import os

def copy_file_with_new_names(input_file, names_list):
    folder = r"C:\Users\acarson\Desktop\JESS Drawings\PDFs"
    # Check if the input file exists
    if not os.path.isfile(input_file):
        print(f"The file {input_file} does not exist.")
        return

    # Loop through the list of names and create copies
    for name in names_list:
        # Create the new file name by adding the extension
        new_file_name = f"{folder}\\{name}.pdf"
        try:
            shutil.copy(input_file, new_file_name)
            print(f"Copied to {new_file_name}")
        except Exception as e:
            print(f"Error copying to {new_file_name}: {e}")

# Example usage
if __name__ == "__main__":
    input_file = r"C:\Users\acarson\Desktop\JESS Drawings\ACCM-GREEN FRAME SINGLE UNIT - 2FAN, TC, NHR, 4EX HEAD.PDF"  # Replace with your input file
    names_list = ["ACCA-GREEN FRAME SINGLE UNIT - 1FAN, SC, NHR, 4EX HEAD","ACCA-GREEN FRAME SINGLE UNIT - 1FAN, TC, NHR, 4EX HEAD",
                  "ACCA-GREEN FRAME SINGLE UNIT - 1FAN, VSSC, NHR, 4EX HEAD","ACCA-GREEN FRAME SINGLE UNIT - 2FAN, SC, NHR, 4EX HEAD",
                  "ACCA-GREEN FRAME SINGLE UNIT - 2FAN, TC, NHR, 4EX HEAD","ACCA-GREEN FRAME SINGLE UNIT - 2FAN, VSSC, NHR, 4EX HEAD",
                  "ACCA-GREEN FRAME SINGLE UNIT - 1FAN, SC, HR, 4 HEAD","ACCA-GREEN FRAME SINGLE UNIT - 1FAN, VSSC, HR, 4EX HEAD",
                  "ACCA-GREEN FRAME SINGLE UNIT - 2FAN, SC, HR, 4 HEAD","ACCA-GREEN FRAME SINGLE UNIT - 2FAN, VSSC, HR, 4EX HEAD",
                  ]  # Replace with your desired names
    copy_file_with_new_names(input_file, names_list)