import os
import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog
import traceback

# Create a Tkinter root window to open a file dialog and select a directory
root = tk.Tk()
root.withdraw()

# Open a file dialog to select a directory
directory = filedialog.askdirectory(title="Select a directory")

# Create a nested dictionary to store the processed codes grouped by code length
all_codes_dict = {1: {}, 2: {}, 3: {}, 4: {}}

# Create a list to store the filenames of files that failed to process
failed_files = []

# Create a variable to keep track of the user's choice to skip lines for all files
skip_lines_all = False

# Iterate over each file in the directory
for filename in os.listdir(directory):
    # Check if the file is an Excel file
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        filepath = os.path.join(directory, filename)

        try:
            # Ask user if they want to skip the first two lines
            if not skip_lines_all:
                skip_lines = tk.messagebox.askyesnocancel("Skip Lines", f"Skip first two lines of {filename}?")

                if skip_lines is None:  # user clicked Cancel
                    break
                elif skip_lines:  # user clicked Yes
                    skip_lines_all = tk.messagebox.askyesno("Skip Lines", "Skip first two lines for all files?")
            else:
                skip_lines = True

            # Load Excel file into a pandas dataframe
            if skip_lines:
                df = pd.read_excel(filepath, sheet_name=0, usecols=["Code", "Title"], skiprows=2)
            else:
                df = pd.read_excel(filepath, sheet_name=0, usecols=["Code", "Title"])

            # Iterate over each row in the dataframe
            for index, row in df.iterrows():
                code = row["Code"]
                default_name = row["Title"]

                # Create a dictionary with the "defaultName" and "value" properties
                code_dict = {"defaultName": default_name, "value": code}

                # Determine the code length and add the code to the appropriate group in the nested dictionaries
                code_length = len(code.split("_"))
                if code_length in all_codes_dict:
                    all_codes_dict[code_length][code] = code_dict

        except Exception as e:
            # If an exception occurs, add the file to the failed files list and continue with the next file
            failed_files.append(filename)
            traceback.print_exc()
            continue

# Save each group of codes to a separate JSON file with the header and footer
for code_length, codes_dict in all_codes_dict.items():
    if codes_dict:
        # Convert the nested dictionary to a list of dictionaries
        codes_list = list(codes_dict.values())

        # Create a new dictionary to store the header and footer lists
        header_footer_dict = {
            "hashedProjectId": "Please enter",
            "attributeTypeId": 5,
            "attributeName": "Please enter",
            "description": "Please enter",
            "inputTypeId": 26,
            "inputValueList": codes_list,
            "inputTypeName": "Dropdown list",
            "valueRequired": True
        }

        # Save each group of codes to a separate JSON file with the header and footer
for code_length, codes_dict in all_codes_dict.items():
    if codes_dict:
        # Convert the nested dictionary to a list of dictionaries
        codes_list = list(codes_dict.values())

        # Create a new dictionary to store the header and footer lists
        header_footer_dict = {
            "hashedProjectId": "Please enter",
            "attributeTypeId": 5,
            "attributeName": "Please enter",
            "description": "Please enter",
            "inputTypeId": 26,
            "inputValueList": codes_list,
            "inputTypeName": "Dropdown list",
            "valueRequired": True
        }

        # Save the processed codes list to a JSON file with the appropriate name
        output_filename = f"Level{code_length}.json"
        output_filepath = os.path.join(directory, output_filename)
        with open(output_filepath, "w") as json_file:
            json_file.write("[")
            json.dump(header_footer_dict, json_file, separators=(',', ':'), indent=2)
            json_file.write("]")

# Add a , character and a newline to the end of each JSON file except the last one
json_filenames = os.listdir(directory)
json_filenames.sort()
for i in range(len(json_filenames) - 1):
    filename = json_filenames[i]
    filepath = os.path.join(directory, filename)
    with open(filepath, "a") as json_file:
        json_file.write("\n")

# Add a closing square bracket to the end of the last JSON file
if json_filenames:
    last_json_filename = json_filenames[-1]
    last_json_filepath = os.path.join(directory, last_json_filename)
    with open(last_json_filepath, "a") as last_json_file:
        last_json_file.write("]")
