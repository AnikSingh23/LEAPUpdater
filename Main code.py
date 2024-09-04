from win32com.client import Dispatch
import re
# from YearCode import orig_first_col, orig_last_new_col, new_first_col, new_last_col, new_first_year
from YearCode import checkvalues
from Variables import relative_paths_no_year_nrcan

# from YearCode import last_col

L = Dispatch('LEAP.LEAPApplication')
p = L.Branch("\Key Assumptions")



Prov = ["ab", "atl", "bc", "can", "mb", "nb", "nl", "ns", "on", "pe", "qc", "sk", "ter"]
Sec = ["agr", "com", "ind", "res", "tran"]
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
        table = match.group(2).strip()
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
            update_expression(child, OrigFile)
        # Else print an error message
        else:
            print("Unexpected leap file type")


# This function updates the expression using a simple replace feature.
def update_expression(branch, OrigFile):
    vari = branch.Variables("Key Assumption")
    print(branch.Name, "  ", vari.Expression)
    # Utilize function to strip the expression into parts
    extracted_values = extract_parts(vari.Expression)
    print(extracted_values)
    # If information can be extracted change the expression if not leave the expression as is.
    if extracted_values:
        filename, table, ex_first_col, ex_row1, ex_last_col, ex_row2 = extracted_values
        print(filename)
        # If filename is in the list of possible filenames then adjust the equation
        if filename in OrigFile or relative_paths_no_year_nrcan:
            checked = checkvalues(filename, table)
            orig_first_col, orig_last_new_col, new_first_col, new_last_col, new_first_year = checked

            print(filename, ",", table, "!", ex_first_col, ex_row1, ":", ex_last_col, ex_row1, ", ", table, "!", ex_first_col, ex_row2, ":", ex_last_col, ex_row2, sep='')
            filename = "/NRCAN/" + filename
            # Create first expression from original values
            exp = "Interp({},{}!{}{}:{}{},{}!{}{}:{}{})".format(filename, table, orig_first_col, ex_row1, orig_last_new_col, ex_row1, table, orig_first_col, ex_row2, orig_last_new_col, ex_row2)
            # Modify specific portions of the code (the rows moved down one and the file name needs to change)
            ex_row3 = int(ex_row1) + 1
            ex_row4 = int(ex_row2) + 1
            filename2 = filename.replace(".xlsx", " 2019.xlsx")
            # Create Second Expression from new values
            exp2 = "Interp({},{}!{}{}:{}{},{}!{}{}:{}{})".format(filename2, table, new_first_col, ex_row3, new_last_col, ex_row3, table, new_first_col, ex_row4, new_last_col, ex_row4)
            # Create the new expression for LEAP
            vari.Expression = "(Year>" + str(new_first_year) + ")*" + str(exp) + " + (Year<" + str(int(new_first_year) - 1) + ")*" + str(exp2)

    print(branch.Name, "  ", vari.Expression)


# 6 sub categories


update_branch(p, file_paths)
