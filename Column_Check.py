import pandas as pd
import glob
from Variables import temp_folder, YearInput

# Creating two lists to define the Res_Ca file since that is the first file and I assume all others will follow its
# formatting style
Main_Res_File = glob.glob(temp_folder + "\\res*ca*")
Imp_Files = glob.glob(temp_folder + "\\*imp*")
Main_File = []

# Remove any imputed files from the list so the list only contains a single file (or at least should)
for i in Imp_Files:
    #  Note this requires using a try statement because if the same file was entered twice into the list it will try to
    #  delete it twice and throw a value error
    try:
        Main_Res_File.remove(i)
    except ValueError:
        pass
# Selects only the current residential file from the previously renamed files
for i in Main_Res_File:
    # Ensures the file being selected is the most up to date (This value is from the earlier year check done in
    # the Variable.py)
    if YearInput[1:] in i:
        Main_File.append(i)

# rare occurrence but if trying to run a program which does not download all the files so there is no res_ca_e file
if not Main_File:
    Main_File = glob.glob(temp_folder + "\\*.xls")


# Read the file for the years column
dfcheck = pd.read_excel(Main_File[0], sheet_name="Table 1", skiprows=10, nrows=0)
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
first_year = min(year_list[2:])
last_year = max(year_list[2:])

# Finding the corresponding letter for the first and last year using the alphabetical list created earlier
first_col = alphabetical_list[year_list.index(first_year)].upper()
last_col = alphabetical_list[year_list.index(last_year)].upper()
