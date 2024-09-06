import os
import pandas
import glob
import re




# Comment/uncomment the respective source folder line
# Only uncomment one
# source_folder = os.path.dirname(sys.executable)  # for PyInstaller
# source_folder = os.path.dirname(__file__)  # for PyCharm
source_folder = r'C:\Users\anik1\Desktop\Work\LEAP\Test Bed' # for a specific folder



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