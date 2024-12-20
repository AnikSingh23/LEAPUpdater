import os
import pandas
import glob
import re




# Comment/uncomment the respective source folder line
# Only uncomment one
# source_folder = os.path.dirname(sys.executable)  # for PyInstaller
# source_folder = os.path.dirname(__file__)  # for PyCharm
source_folder = r'C:\Users\anik1\Desktop\Work\LEAP\2024_11_21_LEAP_Canada' # for a specific folder
expression = False  #True for expressions and False for direct values.
use_high = False # True for end column being ZZ and False for end column being Dynamic
entire_csv_filename = "Entire variables.csv"
Energy_total_csv_filename = "Energy Total.csv"


folder_path = source_folder
# folder_path = os.path.join(source_folder, "nrcan")

# List files with names containing 2019
files_2019 = glob.glob(os.path.join(folder_path, "*", "*2019*.*"))

# List files without an end year (files without numeric characters in their names)
files_no_year = glob.glob(os.path.join(folder_path, "*", "*[!0-9]*.*"))

# Filter out files with numbers in their names from the no_year list
files_no_year = [file for file in files_no_year if not any(char.isdigit() for char in os.path.basename(file))]

# Print the files for 2019
print("Files for 2019:")
for file in files_2019:
    print(file)

# Print the files for noyear
print("\nFiles for noyear:")
for file in files_no_year:
    print(file)

# Create two lists for relative and absolute paths for 2019 and noyear separately
relative_paths_2019 = [os.path.relpath(file, folder_path) for file in files_2019]
absolute_paths_2019 = files_2019

relative_paths_no_year = [os.path.relpath(file, folder_path) for file in files_no_year]
absolute_paths_no_year = files_no_year

# Define the set of keywords to search for in the file names
Sec = ["agr", "com", "ind", "res", "tra"]

# Keywords to exclude from the filenames
exclude_keywords = ["hydrogen", "dis", "template", "building", "OTHER", "old", "crude", "Uranium"]

# List files without an end year (files without numeric characters in their names)
files_no_year = glob.glob(os.path.join(folder_path, "*", "*[!0-9]*.*"))

# Filter out files with numbers in their names from the no_year list
files_no_year = [file for file in files_no_year if not any(char.isdigit() for char in os.path.basename(file))]

# Create a new list to hold the filtered files that contain any of the Sec values and exclude "hydrogen" or "dis"
filtered_files_sec = []

# Loop through each file in files_no_year
for file in files_no_year:
    file_name = os.path.basename(file).lower()  # Convert the file name to lowercase to handle case sensitivity

    # Check if the file contains any of the keywords from Sec and does not contain exclude_keywords
    if any(sec in file_name for sec in Sec) and not any(exclude in file_name for exclude in exclude_keywords):
        filtered_files_sec.append(file)

# Print the original noyear list of files (just for comparison)
print("\nOriginal noyear Files (no filtering applied):")
for file in files_no_year:
    print(file)

# Print the filtered list of files that contain one of the Sec values and exclude the unwanted keywords
print("\nFiles containing values from Sec (['agr', 'com', 'ind', 'res', 'tra']) but excluding 'hydrogen' and 'dis-':")
for file in filtered_files_sec:
    print(file)

# You can also create relative paths for the filtered files if needed, without affecting the original list
relative_paths_sec = [os.path.relpath(file, folder_path) for file in filtered_files_sec]

# Print the relative paths for the filtered files
print("\nRelative Paths for files containing values from Sec and excluding 'hydrogen' and 'dis-':")
for file in relative_paths_sec:
    print(file)

# Print the lists of files for 2019 (relative paths)
print("\nRelative Paths for 2019:")
for file in relative_paths_2019:
    print(file)

# Print the lists of files for 2019 (absolute paths)
print("\nAbsolute Paths for 2019:")
for file in absolute_paths_2019:
    print(file)

# Print the lists of files for noyear (relative paths)
print("\nRelative Paths for noyear:")
for file in relative_paths_no_year:
    print(file)

# Print the lists of files for noyear (absolute paths)
print("\nAbsolute Paths for noyear:")
for file in absolute_paths_no_year:
    print(file)

relative_paths_no_year_nrcan = []

for entry in relative_paths_no_year:
    new_entry = "NRCAN\\" + entry
    relative_paths_no_year_nrcan.append(new_entry)