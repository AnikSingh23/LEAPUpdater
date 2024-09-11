from numpy.lib.utils import source
from pywin.Demos.app.customprint import PRINTDLGORD
from win32com.client import Dispatch
import re
# from YearCode import orig_first_col, orig_last_new_col, new_first_col, new_last_col, new_first_year
from YearCode import checkvalues
from Variables import relative_paths_no_year_nrcan, source_folder, expression, relative_paths_sec, entire_csv_filename, Energy_total_csv_filename
import pandas as pd
import os
import string
from pathlib import Path

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


# Create a function to turn column letters to numbers
def excel_column_to_number(col_str):
    """
    Convert an Excel-style column (e.g., 'A', 'B', 'AA') to a 0-based index (0 for 'A', 1 for 'B', etc.).
    """
    col_str = col_str.upper()  # Ensure the input is uppercase to handle lowercase inputs
    num = 0
    for i, char in enumerate(reversed(col_str)):
        num += (string.ascii_uppercase.index(char) + 1) * (26 ** i)
    return num - 1  # Subtract 1 to convert to 0-based indexing


# Create a function to extract expression information from a pattern
def extract_parts_interp(input_string):
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
# Create a function to extract expression information from a pattern
def extract_parts_data(input_string):
    # Define a regular expression pattern to match the desired parts "basic form is Interp(ab\ab com.xlsx,table 1!c10:l10,table 1!c14:l14)"
    pattern = r'Data\(([^,]+),\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+),\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)\)'

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
def update_branch(branch, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total):
    for child in branch.Children:
        # If branch is another branch go a step deeper
        if child.BranchType == 9:
            print("category", "  ", child.Name)
            update_branch(child, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total)
        # If branch contains an expression modify the expression
        elif child.BranchType == 10:
            print("Key assumption", child.Name)
            if expression:
                update_expression_with_exp(child, OrigFile)
            if not expression:
                update_expression_with_value(child, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total)
                first_write, first_write_total = update_expression_with_value(child, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total)

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
    extracted_values = extract_parts_interp(vari.Expression)
    if not extracted_values:
        extracted_values = extract_parts_data(vari.Expression)
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
                print(filename)
                checked = checkvalues(filename, table)
                orig_first_col, orig_last_new_col = checked

                print(filename, ",", table, "!", ex_first_col, ex_row1, ":", ex_last_col, ex_row1, ", ", table, "!", ex_first_col, ex_row2, ":", ex_last_col, ex_row2, sep='')

                # Create first expression from original values
                exp = "Interp({},{}!{}{}:{}{},{}!{}{}:{}{})".format(filename, table, orig_first_col, ex_row1, orig_last_new_col, ex_row1, table, orig_first_col, ex_row2, orig_last_new_col, ex_row2)

                vari.Expression = exp

    print(branch.Name, "  ", vari.Expression)


def update_expression_with_value(branch, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total):
    vari = branch.Variables("Key Assumption")
    print(branch.Name, "  ", vari.Expression)
    name = vari.Expression

    # Utilize function to strip the expression into parts
    extracted_values = extract_parts_interp(vari.Expression)
    if not extracted_values:
        extracted_values = extract_parts_data(vari.Expression)
    print(extracted_values)

    if extracted_values:
        # Extract relevant details from the expression
        filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2 = extracted_values

        print(f"Filename: {filename}, Table: {table}")
        print(filename)
        print(relative_paths_sec)

        # Check if the table is in the "Table #" or "Table ##" format
        if not re.match(r"^Table \d{1,2}$", table):
            print(f"Skipping update for table {table}, as it doesn't match 'Table #' or 'Table ##' format.")
        else:
            # Check if filename matches
            # Normalize filename
            normalized_filename = Path(filename).as_posix().lower()

            # Check if normalized filename is in the normalized relative_paths_sec
            if any(normalized_filename == Path(path).as_posix().lower() for path in relative_paths_sec):
                # Read the Excel file to get the values
                print(source_folder)
                print(filename)
                absolute_filename = source_folder + "\\" + filename
                df = pd.read_excel(absolute_filename, sheet_name=table)

                checked = checkvalues(filename, table)
                orig_first_col, orig_last_col = checked

                print(ex_row1)
                print("----")
                print(ex_first_col)
                print(orig_last_col)

                ex_first_col_num = excel_column_to_number(ex_first_col)
                ex_last_col_num = excel_column_to_number(orig_last_col) + 1

                print(ex_first_col_num)
                print(ex_last_col_num)
                print("---")

                # Make sure ex_row1 is an integer (if it's coming as a string, convert it) Also minus 2 from the rows as that seems necessary
                ex_row1 = int(ex_row1) - 2  # Convert row index to an integer
                ex_row2 = int(ex_row2) - 2  # Convert row index to an integer (if necessary)

                # Extract rows: row1 contains the years, row2 contains the corresponding values
                years = df.iloc[ex_row1, ex_first_col_num:ex_last_col_num].values.tolist()  # Use .values.tolist() for a Series
                values = df.iloc[ex_row2, ex_first_col_num:ex_last_col_num].values.tolist()  # Similarly for the second row

                print(f"Years: {years}")
                print(f"Values: {values}")

                # Construct the 'interp' expression using the extracted years and values
                interp_parts = [f"{years[i]}, {values[i]}" for i in range(len(years))]  # Alternate between years and values
                interp_expression = "interp(" + ", ".join(interp_parts) + ")"

                # Update the expression in the 'vari' variable
                vari.Expression = interp_expression

                # Prepare the row to be written
                row_to_write = [f"{branch.Name},{table},{filename}"] + values

                # Now we handle the CSV writing part
                # Check if this is the first write to the original CSV file
                if first_write:
                    # Write the years row and then the first data row
                    header = ["Years"] + years
                    df_to_write = pd.DataFrame([header, row_to_write])
                    df_to_write.to_csv(csv_filename, mode='w', index=False, header=False)  # Write header and the first row
                    first_write = False  # After first write, set this to False
                else:
                    # Append the new row without the header
                    df_to_write = pd.DataFrame([row_to_write])
                    df_to_write.to_csv(csv_filename, mode='a', index=False, header=False)  # Append without header

                # If 'total' is in the branch name, write to a separate CSV
                if "total" in branch.Name.lower():
                    # Check if this is the first write to the total CSV file
                    if first_write_total:
                        # Write the years row and then the first data row
                        header = ["Years"] + years
                        df_to_write_total = pd.DataFrame([header, row_to_write])
                        df_to_write_total.to_csv(total_csv_filename, mode='w', index=False, header=False)  # Write header and the first row
                        first_write_total = False
                    else:
                        # Append the new row without the header
                        df_to_write_total = pd.DataFrame([row_to_write])
                        df_to_write_total.to_csv(total_csv_filename, mode='a', index=False, header=False)  # Append without header

    print(branch.Name, "  ", vari.Expression)
    return first_write, first_write_total  # Return updated value of first_write


# 6 sub categories

print()
first_write_total = True
first_write = True
update_branch(p, file_paths, entire_csv_filename, Energy_total_csv_filename, first_write, first_write_total)
