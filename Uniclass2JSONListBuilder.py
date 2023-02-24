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

            # Save the processed codes to a JSON file with the same name as the input file
            output_filename = os.path.splitext(filename)[0] + "_List.json"
            output_filepath = os.path.join(directory, output_filename)
            with open(output_filepath, "w") as json_file:
                json.dump(list(all_codes_dict.values()), json_file, separators=(',', ':'), indent=2)

        except Exception as e:
            # If there was an error processing the file, add it to the list of failed files
            failed_files.append(filename)
            traceback.print_exc()
            continue

# If there were failed files, show a message box with the list of failed files
if failed_files:
    tk.messagebox.showwarning("Failed Files", f"The following files failed to process: {', '.join(failed_files)}")

# Save each group of codes to a separate JSON file
for code_length, codes_dict in all_codes_dict.items():
    if codes_dict:
        # Convert the nested dictionary to a list of dictionaries
        codes_list = list(codes_dict.values())

        # Save the processed codes list to a JSON file with the appropriate name
        output_filename = f"Level{code_length}.json"
        output_filepath = os.path.join(directory, output_filename)
        with open(output_filepath, "w") as json_file:
            json.dump(codes_list, json_file, separators=(',', ':'), indent=2)

