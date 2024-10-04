import pandas as pd
import os
from pathlib import Path
import string
from win32com.client import Dispatch

L = Dispatch('LEAP.LEAPApplication')


# Function to convert Excel-style columns to numbers
def excel_column_to_number(col_str):
    col_str = col_str.upper()
    num = 0
    for i, char in enumerate(reversed(col_str)):
        num += (string.ascii_uppercase.index(char) + 1) * (26 ** i)
    return num - 1


# Check if CSV exists and load it if it does
def load_existing_csv(csv_filename):
    if Path(csv_filename).is_file():
        return pd.read_csv(csv_filename)
    return None


# Extract parts from an expression using regex for interpolation
def extract_parts_interp(input_string):
    pattern = r'Interp\(([^,]+),\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)\s*,\s*([^!]+)!([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)\)'
    matches = list(re.finditer(pattern, input_string))
    if matches:
        match = matches[0]
        return match.group(1), match.group(2), match.group(3), match.group(4), match.group(5), match.group(6)
    return None


# Add row numbers to CSV and handle writing
def write_to_csv(csv_filename, data, headers, first_write=False):
    mode = 'w' if first_write else 'a'
    df = pd.DataFrame([data], columns=headers)

    if first_write:
        df.to_csv(csv_filename, mode=mode, index_label='Row Number')
    else:
        df.to_csv(csv_filename, mode=mode, index_label='Row Number', header=False)


# Function to compare branch path from CSV and update expressions
def compare_and_update_from_csv(branch, csv_filename):
    csv_data = load_existing_csv(csv_filename)
    if csv_data is not None:
        for index, row in csv_data.iterrows():
            csv_branch = row['Branch Path']
            if branch.Name == csv_branch:
                print(f"Match found for branch {branch.Name} in CSV")
                filename = row['Filename']
                table = row['Table']
                ex_first_col = row['First Column']
                ex_row1 = row['Row 1']
                ex_last_col = row['Last Column']
                ex_row2 = row['Row 2']

                # Create new expression based on the CSV
                new_expression = f"Interp({filename},{table}!{ex_first_col}{ex_row1}:{ex_last_col}{ex_row1},{table}!{ex_first_col}{ex_row2}:{ex_last_col}{ex_row2})"
                branch.Variables("Key Assumption").Expression = new_expression
                print(f"Updated expression for {branch.Name} from CSV")


# Function to recursively update branches and compare with CSV
def update_branch(branch, csv_filename, total_csv_filename, first_write, first_write_total):
    for child in branch.Children:
        if child.BranchType == 9:  # If it's a category, recurse
            update_branch(child, csv_filename, total_csv_filename, first_write, first_write_total)
        elif child.BranchType == 10:  # If it's a key assumption, check expression
            print(f"Processing key assumption: {child.Name}")
            if not child.Variables("Key Assumption").Expression:
                compare_and_update_from_csv(child, csv_filename)
            else:
                # Handle new expression or incorrect format
                first_write, first_write_total = update_expression_with_value(child, csv_filename, total_csv_filename, first_write, first_write_total)


# Function to update expressions with values from the Excel files and mark them as "Values"
def update_expression_with_value(branch, csv_filename, total_csv_filename, first_write, first_write_total):
    variable = branch.Variables("Key Assumption")
    expression = variable.Expression

    if expression.startswith("Interp") or expression.startswith("Data"):
        extracted_values = extract_parts_interp(expression)
        if extracted_values:
            filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2 = extracted_values

            # Read the Excel file
            absolute_filename = os.path.join(source_folder, filename)
            df = pd.read_excel(absolute_filename, sheet_name=table)

            ex_first_col_num = excel_column_to_number(ex_first_col)
            ex_last_col_num = excel_column_to_number(ex_last_col) + 1

            # Convert row indexes to integers (adjust for 0-based indexing)
            ex_row1 = int(ex_row1) - 2
            ex_row2 = int(ex_row2) - 2

            # Extract years and values from the Excel file
            years = df.iloc[ex_row1, ex_first_col_num:ex_last_col_num].values.tolist()
            values = df.iloc[ex_row2, ex_first_col_num:ex_last_col_num].values.tolist()

            # Construct the expression using the extracted years and values
            interp_parts = [f"{years[j]}, {values[j]}" for j in range(len(years))]
            new_expression = f"Interp({', '.join(interp_parts)})"
            variable.Expression = new_expression

            # Write the updated data into the CSV files with source type "Value"
            row_to_write = [f"{branch.Name},{table},{filename}, Value"] + values
            write_to_csv(csv_filename, row_to_write, headers=['Branch', 'Table', 'Filename', 'Source Type'] + years, first_write)
            first_write = False

            # Handle writing to the total CSV file
            if "total" in branch.Name.lower() or "aggregate" in branch.Name.lower():
                write_to_csv(total_csv_filename, row_to_write, headers=['Branch', 'Table', 'Filename', 'Source Type'] + years, first_write_total)
                first_write_total = False

    return first_write, first_write_total


# Main processing starts here
first_write = True
first_write_total = True
csv_filename = 'path_to_csv.csv'
total_csv_filename = 'path_to_total_csv.csv'

# Update branches and handle CSV processing
update_branch(L.Branch("Key Assumptions"), csv_filename, total_csv_filename, first_write, first_write_total)
