from numpy.lib.utils import source
from win32com.client import Dispatch
import re
# from YearCode import orig_first_col, orig_last_new_col, new_first_col, new_last_col, new_first_year
from YearCode import checkvalues
from Variables import relative_paths_no_year_nrcan, source_folder, expression, relative_paths_sec
import pandas as pd
import os

# from YearCode import last_col

L = Dispatch('LEAP.LEAPApplication')
p = L.Branch("\Key Assumptions")

Prov = ["ab", "atl", "bc", "can", "mb", "nb", "nl", "ns", "on", "pe", "qc", "sk", "ter"]
Sec = ["agr", "com", "ind", "res", "tra"]
file_paths = []
# Nested for loop to sort through regions
for i in Prov:
    for j in Sec:
        # Creating new file path
        file_path = "\\".join([i, i + " " + j + ".xlsx"])
        file_paths.append(file_path)

print(file_paths)


# Create a function to extract expression information from a pattern
def extract_parts(input_string):
    # Define a regular expression pattern to match the desired parts "basic form is Interp(ab\ab com.xlsx,table 1!c10:l10,table 1!c14:l14)"
    pattern = r'Interp\(([^,]+),\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+),\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)\)'

    # Use re.finditer to find all matches in the input string
    matches = list(re.finditer(pattern, input_string))

    if matches:
        # Extract the matched groups from the first match
        match = matches[0]
        filename = match.group(1).strip()
        table = match.group(2).strip().capitalize()
        ex_first_col = match.group(3).strip()
        ex_row1 = match.group(4).strip()
        ex_last_col = match.group(5).strip()
        ex_row2 = match.group(11).strip()

        return filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2
    else:
        # Return None if no match is found
        return None


# This function will check a branch has children or key assumptions. If it finds another category it will put the new category back into the same function and recursively go through the categories.
# If it finds a key assumption it will pass that assumption to the next function which updates the expression.
def update_branch(branch, OrigFile):
    for child in branch.Children:
        # If branch is another branch go a step deeper
        if child.BranchType == 9:
            print("category", "  ", child.Name)
            update_branch(child, OrigFile)
        # If branch contains an expression modify the expression
        elif child.BranchType == 10:
            print("Key assumption", child.Name)
            if expression:
                update_expression_with_exp(child, OrigFile)
            if not expression:
                update_expression_with_value(child, OrigFile)

        # Else print an error message
        else:
            print("Unexpected leap file type")


# # Helper function to read values from the Excel file
# def read_excel_values(filename, table, first_col, row1, last_col, row2):
#     """
#     Reads the relevant data from the specified Excel file and returns two lists:
#     - years: a list of years
#     - values: a list of corresponding values
#     """
#
#
#     # Load the Excel file into a pandas DataFrame
#     excel_path = os.path.join(source_folder, filename)  # Adjust path as needed
#     print(table)
#     df = pd.read_excel(excel_path, sheet_name=table, header=None)
#
#     # Convert column letters to indices (if columns are provided in letters like 'C', 'L', etc.)
#     def col_letter_to_index(col_letter):
#         return ord(col_letter.upper()) - ord('A')
#
#     # Extract years and corresponding values (adjusting for zero-based indexing)
#     first_col_index = col_letter_to_index(first_col)
#     last_col_index = col_letter_to_index(last_col)
#
#     years = df.iloc[int(row1) - 1:int(row2), first_col_index:last_col_index + 1].index.tolist()
#     values = df.iloc[int(row1) - 1:int(row2), first_col_index:last_col_index + 1].values.tolist()
#
#     return years, values
#
#
# # Main function to update the expression
# def update_expression(branch, OrigFile):
#     vari = branch.Variables("Key Assumption")
#     print(branch.Name, "  ", vari.Expression)
#
#     # Utilize function to strip the expression into parts
#     extracted_values = extract_parts(vari.Expression)
#     print(extracted_values)
#
#     if extracted_values:
#         filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2 = extracted_values
#         print(filename)
#
#         # If filename is in the list of possible filenames then adjust the equation
#         if filename in OrigFile or relative_paths_no_year_nrcan:
#             # Read actual values from the Excel file
#             years, values = read_excel_values(filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2)
#
#             # Create the expression in the format Interp(year1, value1, year2, value2, ...)
#             interp_expression = "Interp(" + ", ".join(f"{year}, {value}" for year, value in zip(years, values)) + ")"
#
#             # Replace the original expression with the new one
#             vari.Expression = interp_expression
#
#     print(branch.Name, "  ", vari.Expression)

# This function updates the expression using a simple replace feature.

def update_expression_with_exp(branch, OrigFile):
    vari = branch.Variables("Key Assumption")
    print(branch.Name, "  ", vari.Expression)

    # Utilize function to strip the expression into parts
    extracted_values = extract_parts(vari.Expression)
    print(extracted_values)

    # If information can be extracted change the expression if not leave the expression as is.
    if extracted_values:
        filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2 = extracted_values
        print(filename)

        # Check if the table is in the "Table #" or "Table ##" format
        if not re.match(r"^Table \d{1,2}$", table):
            print(f"Skipping update for table {table}, as it doesn't match 'Table #' or 'Table ##' format.")
        else:
            # If filename is in the list of possible filenames, then adjust the equation
            if filename in OrigFile and relative_paths_sec:
                checked = checkvalues(filename, table)
                orig_first_col, orig_last_new_col, new_first_col, new_last_col, new_first_year = checked

                print(filename, ",", table, "!", ex_first_col, ex_row1, ":", ex_last_col, ex_row1, ", ", table, "!", ex_first_col, ex_row2, ":", ex_last_col, ex_row2, sep='')

                # Create first expression from original values
                exp = "Interp({},{}!{}{}:{}{},{}!{}{}:{}{})".format(filename, table, orig_first_col, ex_row1, orig_last_new_col, ex_row1, table, orig_first_col, ex_row2, orig_last_new_col, ex_row2)

                vari.Expression = exp

    print(branch.Name, "  ", vari.Expression)


def update_expression_with_value(branch, OrigFile):
    vari = branch.Variables("Key Assumption")
    print(branch.Name, "  ", vari.Expression)

    # Utilize function to strip the expression into parts
    extracted_values = extract_parts(vari.Expression)
    print(extracted_values)

    if extracted_values:
        # Extract relevant details from the expression
        filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2 = extracted_values
        print(f"Filename: {filename}, Table: {table}")

        # Check if filename matches
        if filename in OrigFile or relative_paths_no_year_nrcan:
            # Read the Excel file to get the values
            print(source_folder)
            print(filename)
            absolute_filename = source_folder + "\\" + filename
            df = pd.read_excel(absolute_filename, sheet_name=table)

            # Extract rows: row1 contains the years, row2 contains the corresponding values
            # Adjust for 0-indexing in pandas (Excel uses 1-indexing, so subtract 1 from the row numbers)
            years = df.iloc[ex_row1 - 1, ex_first_col:ex_last_col + 1].tolist()
            values = df.iloc[ex_row2 - 1, ex_first_col:ex_last_col + 1].tolist()

            print(f"Years: {years}")
            print(f"Values: {values}")

            # Construct the 'interp' expression using the extracted years and values
            interp_parts = [f"interp({years[i]}, {values[i]})" for i in range(len(years))]
            interp_expression = ", ".join(interp_parts)

            # Update the expression in the 'vari' variable
            vari.Expression = interp_expression

    print(branch.Name, "  ", vari.Expression)


# 6 sub categories

print()

update_branch(p, file_paths)
