# import pandas as pd
# import openpyxl
# import xlrd
# import pyexcel
# import sys
import glob
import os
import time
from Variables import temp_folder, onedfImp, twodfImp, threedfImp, overwrite, end_variable, misslist, conversion
from Column_Check import first_col, last_col

# Clear missing list, so it does not carry over data from other runs
misslist.clear()

# Define start time of program
GlobalStartTime = time.time()
# defining static variables
filename = []
filenameCa = []
TCOM_File = []
TCOM_Ca_File = []

# search for files with com and .xls in it # If a different format is required these searches can be modified
COM_File = glob.glob(temp_folder + "\\com*.xls")
# search for files with imp in it (so it doesn't try to impute and already imputed file
ExistingComImp = glob.glob(temp_folder + "\\*imp*")
# Add the default canada com file to add to exclusions
ExistingComImp.append(temp_folder + "\\com_ca_e.xls")
# Add any files with ca and .xls to exclusion list
ExistingComImp.extend(glob.glob(temp_folder + "\\*ca*.xls"))

# Removes already imputed files and the canada wide files from the list (first "if" removes any files from exclusion
# list ExistingComImp, the for loop will add any files from the COM_File list which has a ca and .xls in it to its
# own list for further use)
for files in COM_File:
    if files not in ExistingComImp:
        TCOM_File.append(files)
    for s in glob.glob(temp_folder + "\\*ca*"):
        if files == s:
            TCOM_Ca_File.append(files)

# This for loop goes through the filtered files list (this one specifically the files with com and .xls in somewhere
# and removes the CA file and any file with imp in the file name) note this will create/overwrite any imp files by
# creating them (since imp was excluded the for loop creates a new imp path and adds it to a list to be processed
# later. The reason it adds it to a list no matter what is incase if overwriting is not selected the files will be
# checked anyway (I could create a user input to prevent this, but I don't think that is necessary)
for t in TCOM_File:
    file_name = os.path.basename(t)
    file_name_imp = os.path.splitext(file_name)[0] + "_imp.xlsx"
    # Checks user input if files should be overwritten if Y then work as normal if anything else then Y then check if
    # file exists if not create file
    if overwrite == "Y" or overwrite == "y":
        print("working on converting file " + file_name + " to .xlsx format")
        # Pass the original file name and new file name to conversion function (see variables.py for conversion function)
        conversion(file_name, file_name_imp)
        # Creates a list of files for each specific category/style of Excel sheet
        filename.append(file_name_imp)
        print("Finished converting and possibly overwriting old " + file_name_imp)
    # This else if statement checks if the file already exists as an imputed.xlsx file
    elif os.path.splitext(t)[0] + "_imp.xlsx" in ExistingComImp:
        print(file_name_imp + " already exists no need to convert")
        filename.append(file_name_imp)
    # This final else catches all files which do not satisfy the previous conditions
    else:
        print("working on converting file " + file_name + " to .xlsx format")
        conversion(file_name, file_name_imp)
        filename.append(file_name_imp)
        print("Finished creating and converting " + file_name_imp)
print()
for t in TCOM_Ca_File:
    file_name = os.path.basename(t)
    file_name_imp = os.path.splitext(file_name)[0] + "_imp.xlsx"
    # Checks user input if files should be overwritten if Y then work as normal if anything else then Y then check if
    # file exists if not create file
    if overwrite == "Y" or overwrite == "y":
        print("working on converting file " + file_name + " to .xlsx format")
        conversion(file_name, file_name_imp)
        filenameCa.append(file_name_imp)
        print("Finished converting and possibly overwriting old " + file_name_imp)
    # This else if statement checks if the file already exists as an imputed.xlsx file
    elif os.path.splitext(t)[0] + "_imp.xlsx" in ExistingComImp:
        print(file_name_imp + " already exists no need to convert")
        filenameCa.append(file_name_imp)
    # This final else catches all files which do not satisfy the previous conditions
    else:
        print("working on converting file " + file_name + " to .xlsx format")
        conversion(file_name, file_name_imp)
        filenameCa.append(file_name_imp)
        print("Finished creating and converting " + file_name_imp)

# Test function to determine the time of conversion
# Define time at end of program
ConEndTime = time.time()
# Determine the time and convert to minutes and seconds
ConTimeMin, ConTimeSec = divmod((ConEndTime - GlobalStartTime) / 60, 1.0)
print("Conversion completion time: " + str(round(ConTimeMin)) + " Minutes and " + str(round(ConTimeSec * 60)) +
      " Seconds")

# Impute the regular files. What this "for" loop does is procedurally goes through the filename list which was created
# earlier and defines a new variable name for each file which then gets input into the dfimp functions to indicate
# which file it is working on
for name in filename:
    tables1 = ["Table 1", "Table 4", "Table 5", "Table 6", "Table 7", "Table 8", "Table 9", "Table 10", "Table 11",
               "Table 12", "Table 13", "Table 14", "Table 15", "Table 16", "Table 17", "Table 18", "Table 19",
               "Table 20", "Table 21", "Table 22", "Table 23", "Table 24", "Table 26", "Table 28"]
    tables2 = ["Table 2"]
    tables3 = ["Table 3", "Table 25", "Table 27", "Table 29", "Table 30", "Table 31", "Table 33"]
    tables4 = ["Table 32"]
    tables5 = ["Table 34"]
    tables6 = ["Table 35"]
    tables7 = ["Table 37", "Table 41", "Table 43", "Table 45", "Table 47", "Table 51", "Table 53"]
    tables8 = ["Table 36", "Table 38", "Table 40", "Table 42", "Table 44", "Table 46", "Table 48", "Table 50",
               "Table 52", "Table 54"]
    tables9 = ["Table 39"]
    print()
    print("Working on Imputing file " + name)
    print()
    # For loop to impute each table separately
    for x in tables1:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 20, name, x)
        onedfImp(first_col, 31, last_col, 31, name, x)
    for x in tables2:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 21, name, x)
        onedfImp(first_col, 33, last_col, 33, name, x)
    for x in tables3:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 24, name, x)
        onedfImp(first_col, 39, last_col, 39, name, x)
    for x in tables4:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        onedfImp(first_col, 23, last_col, 23, name, x)
    for x in tables5:
        print("Working on " + x)
        onedfImp(first_col, 13, last_col, 13, name, x)
    for x in tables6:
        print("Working on " + x)
        onedfImp(first_col, 13, last_col, 13, name, x)
        onedfImp(first_col, 16, last_col, 16, name, x)
        onedfImp(first_col, 21, last_col, 21, name, x)
        onedfImp(first_col, 24, last_col, 24, name, x)
        twodfImp(first_col, 29, last_col, 29, first_col, 31, last_col, 36, name, x)
        onedfImp(first_col, 47, last_col, 47, name, x)
        twodfImp(first_col, 52, last_col, 52, first_col, 54, last_col, 59, name, x)
        onedfImp(first_col, 70, last_col, 70, name, x)
        # notice for fix later tables 6 and 7 are the exact same data ranges and can be combined
    for x in tables7:
        print("Working on " + x)
        onedfImp(first_col, 13, last_col, 13, name, x)
        onedfImp(first_col, 16, last_col, 16, name, x)
        onedfImp(first_col, 21, last_col, 21, name, x)
        onedfImp(first_col, 24, last_col, 24, name, x)
        twodfImp(first_col, 29, last_col, 29, first_col, 31, last_col, 36, name, x)
        onedfImp(first_col, 47, last_col, 47, name, x)
        twodfImp(first_col, 52, last_col, 52, first_col, 54, last_col, 59, name, x)
        onedfImp(first_col, 70, last_col, 70, name, x)
    for x in tables8:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        onedfImp(first_col, 23, last_col, 23, name, x)
        twodfImp(first_col, 28, last_col, 28, first_col, 30, last_col, 35, name, x)
        onedfImp(first_col, 46, last_col, 46, name, x)
        onedfImp(first_col, 51, last_col, 51, name, x)
        onedfImp(first_col, 52, last_col, 52, name, x)
    for x in tables9:
        print("Working on " + x)
        onedfImp(first_col, 13, last_col, 13, name, x)
        onedfImp(first_col, 16, last_col, 16, name, x)
        onedfImp(first_col, 21, last_col, 21, name, x)
        onedfImp(first_col, 24, last_col, 24, name, x)
        onedfImp(first_col, 29, last_col, 29, name, x)
        onedfImp(first_col, 32, last_col, 32, name, x)
        twodfImp(first_col, 37, last_col, 37, first_col, 39, last_col, 44, name, x)
        onedfImp(first_col, 55, last_col, 55, name, x)

for name in filenameCa:
    tables1 = ["Table 1"]
    tables2 = ["Table 2"]
    tables3 = ["Table 4"]
    tables4 = ["Table 6"]
    tables5 = ["Table 7", "Table 8", "Table 10", "Table 11", "Table 13", "Table 14", "Table 16", "Table 17", "Table 19",
               "Table 20", "Table 22", "Table 23", "Table 25", "Table 26", "Table 28", "Table 29", "Table 31",
               "Table 32", "Table 34", "Table 35", "Table 41", "Table 45"]
    tables6 = ["Table 9", "Table 12", "Table 15", "Table 18", "Table 21", "Table 24", "Table 27", "Table 30",
               "Table 33", "Table 36", "Table 44", "Table 48", "Table 50", "Table 52"]
    tables7 = ["Table 37"]
    tables8 = ["Table 38", "Table 54"]
    tables9 = ["Table 40", "Table 56"]
    tables10 = ["Table 42", "Table 46", "Table 49", "Table 51"]
    tables11 = ["Table 53"]
    tables12 = ["Table 57"]
    tables13 = ["Table 58"]
    tables14 = ["Table 59", "Table 61", "Table 65", "Table 67", "Table 69", "Table 71", "Table 73", "Table 75",
                "Table 77"]
    tables15 = ["Table 60", "Table 62", "Table 64", "Table 66", "Table 68", "Table 70", "Table 72", "Table 74",
                "Table 76", "Table 78"]
    tables16 = ["Table 63"]

    # Tables 3, 5, 39, 43, 47, 55 are GHG tables

    print()
    print("Working on Imputing file " + name)
    print()
    # For loop to impute each table separately
    for x in tables1:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 20, name, x)
        onedfImp(first_col, 31, last_col, 31, name, x)
        onedfImp(first_col, 61, last_col, 61, name, x)
        onedfImp(first_col, 62, last_col, 62, name, x)

    for x in tables2:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 24, name, x)
        onedfImp(first_col, 39, last_col, 39, name, x)
        onedfImp(first_col, 72, last_col, 72, name, x)
        onedfImp(first_col, 73, last_col, 73, name, x)

    for x in tables3:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 21, name, x)
        onedfImp(first_col, 33, last_col, 33, name, x)
        onedfImp(first_col, 59, last_col, 59, name, x)
        onedfImp(first_col, 60, last_col, 60, name, x)

    for x in tables4:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 21, name, x)
        onedfImp(first_col, 33, last_col, 33, name, x)
        onedfImp(first_col, 60, last_col, 60, name, x)
        onedfImp(first_col, 61, last_col, 61, name, x)

    for x in tables5:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 20, name, x)
        onedfImp(first_col, 31, last_col, 31, name, x)

    for x in tables6:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 21, name, x)
        onedfImp(first_col, 33, last_col, 33, name, x)

    for x in tables7:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 20, name, x)
        onedfImp(first_col, 31, last_col, 31, name, x)
        onedfImp(first_col, 61, last_col, 61, name, x)

    for x in tables8:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 24, name, x)
        onedfImp(first_col, 39, last_col, 39, name, x)
        onedfImp(first_col, 72, last_col, 72, name, x)

    for x in tables9:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 21, name, x)
        onedfImp(first_col, 33, last_col, 33, name, x)
        onedfImp(first_col, 60, last_col, 60, name, x)

    for x in tables10:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 24, name, x)
        onedfImp(first_col, 39, last_col, 39, name, x)

    for x in tables11:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        onedfImp(first_col, 23, last_col, 23, name, x)
        onedfImp(first_col, 45, last_col, 45, name, x)

    for x in tables12:
        print("Working on " + x)
        onedfImp(first_col, 13, last_col, 13, name, x)

    for x in tables13:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 21, name, x)

    for x in tables14:
        print("Working on " + x)
        onedfImp(first_col, 13, last_col, 13, name, x)
        onedfImp(first_col, 16, last_col, 16, name, x)
        onedfImp(first_col, 21, last_col, 21, name, x)
        onedfImp(first_col, 24, last_col, 24, name, x)
        twodfImp(first_col, 29, last_col, 29, first_col, 31, last_col, 36, name, x)
        onedfImp(first_col, 47, last_col, 47, name, x)
        twodfImp(first_col, 52, last_col, 52, first_col, 54, last_col, 59, name, x)
        onedfImp(first_col, 70, last_col, 70, name, x)

    for x in tables15:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        onedfImp(first_col, 23, last_col, 23, name, x)
        twodfImp(first_col, 28, last_col, 28, first_col, 30, last_col, 35, name, x)
        onedfImp(first_col, 46, last_col, 46, name, x)
        onedfImp(first_col, 51, last_col, 51, name, x)
        onedfImp(first_col, 52, last_col, 52, name, x)

    for x in tables16:
        print("Working on " + x)
        onedfImp(first_col, 13, last_col, 13, name, x)
        onedfImp(first_col, 16, last_col, 16, name, x)
        onedfImp(first_col, 21, last_col, 21, name, x)
        onedfImp(first_col, 24, last_col, 24, name, x)
        onedfImp(first_col, 29, last_col, 29, name, x)
        onedfImp(first_col, 32, last_col, 32, name, x)
        twodfImp(first_col, 37, last_col, 37, first_col, 39, last_col, 44, name, x)
        onedfImp(first_col, 55, last_col, 55, name, x)

# Define time at end of program
GlobalEndTime = time.time()
# Determine the time and convert to minutes and seconds
COMTimeMin, COMTimeSec = divmod((GlobalEndTime - GlobalStartTime) / 60, 1.0)

print()
print(
    "------------------------------------------------------------------------------------------------------------------------")
# Defining the variable for the amount of missing values
COM_Miss = sum(misslist)
print("There were " + str(COM_Miss) + " missing values overall")

if end_variable:
    print("Overall completion time: " + str(round(COMTimeMin)) + " Minutes and " + str(round(COMTimeSec * 60)) +
          " Seconds")
else:
    print("Commercial and institutional sector completion time: " + str(round(COMTimeMin)) + " Minutes and " + str(
        round(COMTimeSec * 60)) + " Seconds")

print("Hopefully completed imputing commercial and institutional files successfully")
if end_variable:
    input("Press any key to exit console")
else:
    print()
    print(
        "------------------------------------------------------------------------------------------------------------------------")
