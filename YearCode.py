import pandas as pd
import glob
from Variables import source_folder


# Creating two lists to define the Res_Ca file since that is the first file and I assume all others will follow its
# formatting style
# Examined_File = glob.glob(source_folder + "\\*com*")
#
# print(Examined_File)

# def checkvalues(filename, table):
#     try:
#         Examined_File = source_folder + filename
#         print(Examined_File)
#
#         # Read the file for the years column
#         dfcheck = pd.read_excel(Examined_File[1], sheet_name=table, skiprows=9, nrows=0)
#         # Turn column into a list
#         orig_year_list = dfcheck.columns.tolist()
#
#         # Create an alphabetical list the same size as the year list to show corresponding column letters
#         alphabetical_list = []
#         for i in range(len(orig_year_list)):
#             preceding_char_increments = i // 26
#             char_index = i % 26
#             preceding_chars = ''.join([chr(97 + j) for j in range(preceding_char_increments)])
#             alphabetical_list.append(preceding_chars + chr(97 + char_index))
#
#         # Finds the min and max values of the columns excluding the first two
#         print(orig_year_list)
#         orig_first_year = min(orig_year_list[2:])
#         orig_last_year = max(orig_year_list[2:])
#
#         # Find the corresponding letter for the orig_first and orig_last year
#         orig_first_col = alphabetical_list[orig_year_list.index(orig_first_year)].upper()
#         orig_last_col = alphabetical_list[orig_year_list.index(orig_last_year)].upper()
#
#         print(orig_first_year, "   ", orig_first_col)
#         print(orig_last_year, "   ", orig_last_col)
#
#         # Read the file for the years column
#         dfcheck = pd.read_excel(Examined_File[0], sheet_name="Table 1", skiprows=10, nrows=0)
#         year_list = dfcheck.columns.tolist()
#
#         # Create an alphabetical list for the year list
#         alphabetical_list = []
#         for i in range(len(year_list)):
#             preceding_char_increments = i // 26
#             char_index = i % 26
#             preceding_chars = ''.join([chr(97 + j) for j in range(preceding_char_increments)])
#             alphabetical_list.append(preceding_chars + chr(97 + char_index))
#
#         # Finds the min and max values of the columns excluding the first two
#         print(year_list)
#         new_first_year = min(year_list[2:])
#         new_last_year = max(year_list[2:])
#
#         # Find the corresponding letter for the new_first and new_last year
#         new_first_col = alphabetical_list[year_list.index(new_first_year)].upper()
#         new_last_col = alphabetical_list[year_list.index(new_last_year)].upper()
#
#         print(new_first_year, "   ", new_first_col)
#         print(new_last_year, "   ", new_last_col)
#
#         orig_last_new_col = alphabetical_list[orig_year_list.index(new_first_year - 1)].upper()
#         print((new_first_year - 1), "   ", orig_last_new_col)
#
#         return orig_first_col, orig_last_new_col, new_first_col, new_last_col, new_first_year
#
#     except Exception as e:
#         print(f"Error processing file '{filename}': {e}")
#         return None

def checkvalues(filename, table):
    try:
        Examined_File = source_folder + "\\" + filename
        print(Examined_File)

        # Read the file for the years column
        dfcheck = pd.read_excel(Examined_File, sheet_name=table, skiprows=9, nrows=0)
        # Turn column into a list
        orig_year_list = dfcheck.columns.tolist()

        # Create an alphabetical list the same size as the year list to show corresponding column letters
        alphabetical_list = []
        for i in range(len(orig_year_list)):
            preceding_char_increments = i // 26
            char_index = i % 26
            preceding_chars = ''.join([chr(97 + j) for j in range(preceding_char_increments)])
            alphabetical_list.append(preceding_chars + chr(97 + char_index))

        # Finds the min and max values of the columns excluding the first two
        print(orig_year_list)
        orig_first_year = min(orig_year_list[2:])
        orig_last_year = max(orig_year_list[2:])

        # Find the corresponding letter for the orig_first and orig_last year
        orig_first_col = alphabetical_list[orig_year_list.index(orig_first_year)].upper()
        orig_last_col = alphabetical_list[orig_year_list.index(orig_last_year)].upper()

        print(orig_first_year, "   ", orig_first_col)
        print(orig_last_year, "   ", orig_last_col)


        return orig_first_col, orig_last_col

    except Exception as e:
        print(f"Error processing file '{filename}': {e}")
        return None
