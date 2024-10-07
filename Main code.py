from numpy.lib.utils import source
from pywin.Demos.app.customprint import PRINTDLGORD
from win32com.client import Dispatch
import re
# from YearCode import orig_first_col, orig_last_new_col, new_first_col, new_last_col, new_first_year
from YearCode import checkvalues
from Variables import relative_paths_no_year_nrcan, source_folder, expression, relative_paths_sec, entire_csv_filename, Energy_total_csv_filename, use_high
import pandas as pd
import os
import string
from pathlib import Path

# from YearCode import last_col

L = Dispatch('LEAP.LEAPApplication')
p = L.Branch("\Key Assumptions")

Prov = ["ab", "atl", "bc", "bct", "can", "mb", "nb", "nl", "ns", "on", "pe", "qc", "sk", "ter"]
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
    # Define a regular expression pattern to match the desired parts, allowing for extra spaces
    pattern = r'Interp\(([^,]+),\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)\s*,\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)\)'

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
    pattern = r'Data\(([^,]+),\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)\s*,\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)\)'

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
# If it finds a key assumption it will pass that assumption to the next function which updates the expression. 9 = Branch 10 = Expression
def update_branch(branch, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total):
    for child in branch.Children:
        # If branch is another branch go a step deeper
        if child.BranchType == 9:
            print("category", "  ", child.Name)
            # Pass the updated first_write and first_write_total from the recursive call
            first_write, first_write_total = update_branch(child, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total)

        # If branch contains an expression modify the expression
        elif child.BranchType == 10:
            print("Key assumption", child.Name)
            if expression:
                first_write, first_write_total = update_expression_with_exp(child, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total)
            else:
                # Ensure updated values are returned and passed on
                first_write, first_write_total = update_expression_with_value(child, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total)

        # Else print an error message
        else:
            print("Unexpected leap file type")

    # Return updated values to ensure changes propagate through recursion
    return first_write, first_write_total


def update_expression_with_exp(branch, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total, use_zz=False):
    vari = branch.Variables("Key Assumption")
    print(branch.Name, "  ", vari.Expression)

    # Utilize function to strip the expression into parts
    extracted_values = extract_parts_interp(vari.Expression)
    if not extracted_values:
        extracted_values = extract_parts_data(vari.Expression)

    # Check if extracted_values is empty or None
    if not extracted_values:
        print(f"Failed to extract values for expression: {vari.Expression}")
        return first_write, first_write_total  # Exit if extraction fails

    print(extracted_values)

    # Extract relevant details from the expression
    filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2 = extracted_values

    # Check if the table is in the "Table #" or "Table ##" format
    if not re.match(r"^Table \d{1,2}$", table):
        print(f"Skipping update for table {table}, as it doesn't match 'Table #' or 'Table ##' format.")
    else:
        # Normalize filename
        normalized_filename = Path(filename).as_posix().lower()

        # Check if normalized filename is in the normalized relative_paths_sec
        if any(normalized_filename == Path(path).as_posix().lower() for path in relative_paths_sec):
            # Read the Excel file to get the values
            absolute_filename = source_folder + "\\" + filename
            df = pd.read_excel(absolute_filename, sheet_name=table)

            checked = checkvalues(filename, table)
            orig_first_col, orig_last_col = checked

            # Set column to 'ZZ' if use_zz is True, otherwise use the original first and last columns
            if use_high:
                ex_first_col = orig_first_col
                ex_last_col = 'ZZ'
            else:
                ex_first_col = orig_first_col
                ex_last_col = orig_last_col

            ex_first_col_num = excel_column_to_number(ex_first_col)
            ex_last_col_num = excel_column_to_number(ex_last_col) + 1

            # Make sure ex_row1 is an integer (if it's coming as a string, convert it) Also minus 2 from the rows as that seems necessary
            ex_row1 = int(ex_row1) - 2  # Convert row index to an integer
            ex_row2 = int(ex_row2) - 2  # Convert row index to an integer (if necessary)

            # Extract rows: row1 contains the years, row2 contains the corresponding values
            years = df.iloc[ex_row1, ex_first_col_num:ex_last_col_num].values.tolist()  # Use .values.tolist() for a Series
            values = df.iloc[ex_row2, ex_first_col_num:ex_last_col_num].values.tolist()  # Similarly for the second row

            # Construct the original style expression using the extracted parts
            exp = f"Interp({filename},{table}!{ex_first_col}{ex_row1 + 2}:{ex_last_col}{ex_row1 + 2},{table}!{ex_first_col}{ex_row2 + 2}:{ex_last_col}{ex_row2 + 2})"

            # Update the expression in the 'vari' variable
            vari.Expression = exp

            # Prepare the row to be written
            row_to_write = [f"{branch.Name},{table},{filename}"] + values
            print(first_write)
            print(first_write_total)

            # Now we handle the CSV writing part
            if first_write:
                # Write header only for the first time
                header = ["Years"] + years
                df_to_write = pd.DataFrame([header, row_to_write])
                df_to_write.to_csv(csv_filename, mode='w', index=False, header=False)
                first_write = False  # Set to False after first write
            else:
                # Append data in subsequent writes
                print("writing to csv")
                df_to_write = pd.DataFrame([row_to_write])
                df_to_write.to_csv(csv_filename, mode='a', index=False, header=False)

            # Write data to Energy_total_csv_filename
            if "total" in branch.Name.lower() or "end use" in branch.Name.lower() or "aggregate" in branch.Name.lower():
                if first_write_total:
                    header = ["Years"] + years
                    df_to_write_total = pd.DataFrame([header, row_to_write])
                    df_to_write_total.to_csv(total_csv_filename, mode='w', index=False, header=False)
                    first_write_total = False  # Set to False after first write
                else:
                    print("writing to energy csv")
                    df_to_write_total = pd.DataFrame([row_to_write])
                    df_to_write_total.to_csv(total_csv_filename, mode='a', index=False, header=False)

    print(branch.Name, "  ", vari.Expression)
    return first_write, first_write_total  # Return updated value of first_write


def update_expression_with_value(branch, OrigFile, csv_filename, total_csv_filename, first_write, first_write_total):
    vari = branch.Variables("Key Assumption")
    print(branch.Name, "  ", vari.Expression)

    # Check if the expression starts with "interp" or "data"
    if vari.Expression.startswith("Interp"):
        expression_type = "Interp"
    elif vari.Expression.startswith("Data"):
        expression_type = "Data"
    else:
        expression_type = None  # You can handle this case as needed
        print(f"Unrecognized expression type: {vari.Expression}")
        return first_write, first_write_total  # Exit if the expression is not recognized

    # Utilize the correct function to strip the expression into parts based on the flag
    if expression_type == "Interp":
        extracted_values = extract_parts_interp(vari.Expression)
    elif expression_type == "Data":
        extracted_values = extract_parts_data(vari.Expression)

    # Check if extracted_values is empty or None
    if not extracted_values:
        print(f"Failed to extract values for expression: {vari.Expression}")
        return first_write, first_write_total  # Exit if extraction fails

    print(extracted_values)

    if extracted_values:
        # Extract relevant details from the expression
        filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2 = extracted_values

        print(f"Filename: {filename}, Table: {table}")
        print(filename)

        # Check if the table is in the "Table #" or "Table ##" format
        if not re.match(r"^Table \d{1,2}$", table):
            print(f"Skipping update for table {table}, as it doesn't match 'Table #' or 'Table ##' format.")
        else:
            # Normalize filename
            normalized_filename = Path(filename).as_posix().lower()

            # Check if normalized filename is in the normalized relative_paths_sec
            if any(normalized_filename == Path(path).as_posix().lower() for path in relative_paths_sec):
                # Read the Excel file to get the values

                absolute_filename = source_folder + "\\" + filename
                df = pd.read_excel(absolute_filename, sheet_name=table)

                checked = checkvalues(filename, table)
                orig_first_col, orig_last_col = checked

                ex_first_col_num = excel_column_to_number(ex_first_col)
                ex_last_col_num = excel_column_to_number(orig_last_col) + 1

                # Make sure ex_row1 is an integer (if it's coming as a string, convert it) Also minus 2 from the rows as that seems necessary
                ex_row1 = int(ex_row1) - 2  # Convert row index to an integer
                ex_row2 = int(ex_row2) - 2  # Convert row index to an integer (if necessary)

                # Extract rows: row1 contains the years, row2 contains the corresponding values
                years = df.iloc[ex_row1, ex_first_col_num:ex_last_col_num].values.tolist()  # Use .values.tolist() for a Series
                values = df.iloc[ex_row2, ex_first_col_num:ex_last_col_num].values.tolist()  # Similarly for the second row

                # Construct the expression (either 'interp' or 'data') using the extracted years and values
                interp_parts = [f"{years[j]}, {values[j]}" for j in range(len(years))]  # Alternate between years and values
                interp_expression = f"{expression_type}(" + ", ".join(interp_parts) + ")"

                # Update the expression in the 'vari' variable
                vari.Expression = interp_expression

                # Prepare the row to be written
                row_to_write = [f"{branch.Name},{table},{filename}"] + values
                print(first_write)
                print(first_write_total)

                # Now we handle the CSV writing part
                if first_write:
                    # Write header only for the first time
                    header = ["Years"] + years
                    df_to_write = pd.DataFrame([header, row_to_write])
                    df_to_write.to_csv(csv_filename, mode='w', index=False, header=False)
                    first_write = False  # Set to False after first write
                else:
                    # Append data in subsequent writes
                    print("writing to csv")
                    df_to_write = pd.DataFrame([row_to_write])
                    df_to_write.to_csv(csv_filename, mode='a', index=False, header=False)

                # Write data to Energy_total_csv_filename
                if "total" in branch.Name.lower() or "end use" in branch.Name.lower() or "aggregate" in branch.Name.lower():
                    if first_write_total:
                        header = ["Years"] + years
                        df_to_write_total = pd.DataFrame([header, row_to_write])
                        df_to_write_total.to_csv(total_csv_filename, mode='w', index=False, header=False)
                        first_write_total = False  # Set to False after first write
                    else:
                        print("writing to energy csv")
                        df_to_write_total = pd.DataFrame([row_to_write])
                        df_to_write_total.to_csv(total_csv_filename, mode='a', index=False, header=False)

    print(branch.Name, "  ", vari.Expression)
    return first_write, first_write_total  # Return updated value of first_write


# 6 sub categories

print()
global_first_write_total = True
global_first_write = True
global_first_write, global_first_write_total = update_branch(p, file_paths, entire_csv_filename, Energy_total_csv_filename, global_first_write, global_first_write_total)

