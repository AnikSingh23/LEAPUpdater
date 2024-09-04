import pandas as pd
import glob
from Variables import source_folder


# Creating two lists to define the Res_Ca file since that is the first file and I assume all others will follow its
# formatting style
# Examined_File = glob.glob(source_folder + "\\*com*")
#
# print(Examined_File)

def checkvalues(filename, table):
    Examined_File = source_folder + filename

    # Read the file for the years column
    dfcheck = pd.read_excel(Examined_File[1], sheet_name=table, skiprows=9, nrows=0)
    # Turn column into a list
    orig_year_list = dfcheck.columns.tolist()
    # Create an alphabetical list the same size as the year list to show corresponding column letters
    alphabetical_list = []
    # This for loop will populate the alphabetical list with the letters if the number of columns surpass 26 (a-z) then this
    # loop will add a preceding letter appropriately for example column 26 would be Z and 27 would be AA, column 52 would be
    # AZ and column 53 would be BA
    for i in range(len(orig_year_list)):
        # Calculate the number of times the preceding character needs to be incremented
        preceding_char_increments = i // 26
        # Calculate the index of the current character in the alphabet (0-based)
        char_index = i % 26
        # Create the preceding characters by incrementing the character 'a' the number of times calculated
        preceding_chars = ''.join([chr(97 + j) for j in range(preceding_char_increments)])
        # Append the preceding characters and the current character to the result list
        alphabetical_list.append(preceding_chars + chr(97 + char_index))

    # Finds the min and max values of the columns excluding the first two (using 2:) since they are strings and are there
    # for formatting the Excel table
    print(orig_year_list)
    orig_first_year = min(orig_year_list[2:])
    orig_last_year = max(orig_year_list[2:])

    # Finding the corresponding letter for the orig_first and orig_last year using the alphabetical list created earlier
    orig_first_col = alphabetical_list[orig_year_list.index(orig_first_year)].upper()
    orig_last_col = alphabetical_list[orig_year_list.index(orig_last_year)].upper()

    print(orig_first_year, "   ", orig_first_col)
    print(orig_last_year, "   ", orig_last_col)

    # Read the file for the years column
    dfcheck = pd.read_excel(Examined_File[0], sheet_name="Table 1", skiprows=10, nrows=0)
    # Turn column into a list
    year_list = dfcheck.columns.tolist()
    # Create an alphabetical list the same size as the year list to show corresponding column letters
    alphabetical_list = []
    # This for loop will populate the alphabetical list with the letters if the number of columns surpass 26 (a-z) then this
    # loop will add a preceding letter appropriately for example column 26 would be Z and 27 would be AA, column 52 would be
    # AZ and column 53 would be BA
    for i in range(len(year_list)):
        # Calculate the number of times the preceding character needs to be incremented
        preceding_char_increments = i // 26
        # Calculate the index of the current character in the alphabet (0-based)
        char_index = i % 26
        # Create the preceding characters by incrementing the character 'a' the number of times calculated
        preceding_chars = ''.join([chr(97 + j) for j in range(preceding_char_increments)])
        # Append the preceding characters and the current character to the result list
        alphabetical_list.append(preceding_chars + chr(97 + char_index))

    # Finds the min and max values of the columns excluding the first two (using 2:) since they are strings and are there
    # for formatting the Excel table
    print(year_list)
    new_first_year = min(year_list[2:])
    new_last_year = max(year_list[2:])

    # Finding the corresponding letter for the new_first and new_last year using the alphabetical list created earlier
    new_first_col = alphabetical_list[year_list.index(new_first_year)].upper()
    new_last_col = alphabetical_list[year_list.index(new_last_year)].upper()

    print(new_first_year, "   ", new_first_col)
    print(new_last_year, "   ", new_last_col)

    orig_last_new_col = alphabetical_list[orig_year_list.index(new_first_year - 1)].upper()
    print((new_first_year - 1), "   ", orig_last_new_col)

    return orig_first_col, orig_last_new_col, new_first_col, new_last_col, new_first_year
