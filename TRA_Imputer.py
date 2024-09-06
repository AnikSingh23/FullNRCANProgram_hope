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
TTRAN_File = []
TTRAN_Ca_File = []

# search for files with tran and .xls in it # If a different format is required these searches can be modified
TRAN_File = glob.glob(temp_folder + "\\tran*.xls")
# search for files with imp in it (so it doesn't try to impute and already imputed file
ExistingTranImp = glob.glob(temp_folder + "\\*imp*")
# Add the default canada tran file to add to exclusions
ExistingTranImp.append(temp_folder + "\\tran_ca_e.xls")
# Add any files with ca and .xls to exclusion list
ExistingTranImp.extend(glob.glob(temp_folder + "\\*ca*.xls"))

# Removes already imputed files and the canada wide files from the list (first "if" removes any files from exclusion
# list ExistingTranImp, the for loop will add any files from the TRAN_File list which has a ca and .xls in it to its
# own list for further use)
for files in TRAN_File:
    if files not in ExistingTranImp:
        TTRAN_File.append(files)
    for s in glob.glob(temp_folder + "\\*ca*"):
        if files == s:
            TTRAN_Ca_File.append(files)

# This for loop goes through the filtered files list (this one specifically the files with tran and .xls in somewhere
# and removes the CA file and any file with imp in the file name) note this will create/overwrite any imp files by
# creating them (since imp was excluded the for loop creates a new imp path and adds it to a list to be processed
# later. The reason it adds it to a list no matter what is incase if overwriting is not selected the files will be
# checked anyway (I could create a user input to prevent this, but I don't think that is necessary)
for t in TTRAN_File:
    file_name = os.path.basename(t)
    file_name_imp = os.path.splitext(file_name)[0] + "_imp.xlsx"
    file_name_imp = file_name_imp.replace("tran", "tra")
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
    elif os.path.splitext(t)[0] + "_imp.xlsx" in ExistingTranImp:
        print(file_name_imp + " already exists no need to convert")
        filename.append(file_name_imp)
    # This final else catches all files which do not satisfy the previous conditions
    else:
        print("working on converting file " + file_name + " to .xlsx format")
        conversion(file_name, file_name_imp)
        filename.append(file_name_imp)
        print("Finished creating and converting " + file_name_imp)

print()

for t in TTRAN_Ca_File:
    file_name = os.path.basename(t)
    file_name_imp = os.path.splitext(file_name)[0] + "_imp.xlsx"
    file_name_imp = file_name_imp.replace("tran", "tra")
    # Checks user input if files should be overwritten if Y then work as normal if anything else then Y then check if
    # file exists if not create file
    if overwrite == "Y" or overwrite == "y":
        print("working on converting file " + file_name + " to .xlsx format")
        conversion(file_name, file_name_imp)
        filenameCa.append(file_name_imp)
        print("Finished converting and possibly overwriting old " + file_name_imp)
    # This else if statement checks if the file already exists as an imputed.xlsx file
    elif os.path.splitext(t)[0] + "_imp.xlsx" in ExistingTranImp:
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
    # define table list which is used in for loop
    tables1 = ["Table 1"]
    tables2 = ["Table 2"]
    tables3 = ["Table 3"]
    tables4 = ["Table 7"]
    tables5 = ["Table 9"]
    tables6 = ["Table 10", "Table 22", "Table 29"]
    tables7 = ["Table 11", "Table 20", "Table 25", "Table 26", "Table 28", "Table 34", "Table 35"]
    tables8 = ["Table 12"]
    tables9 = ["Table 13"]
    tables10 = ["Table 14", "Table 15", "Table 19"]
    tables11 = ["Table 16"]
    tables12 = ["Table 17", "Table 18"]
    tables13 = ["Table 23"]
    tables14 = ["Table 24", "Table 33"]
    tables15 = ["Table 27"]
    tables16 = ["Table 30", "Table 36"]
    tables17 = ["Table 32"]
    tables18 = ["Table 21"]
    tables19 = ["Table 31"]
    tables20 = ["Table 37"]
    # Ignore table 4,5,6,8 due to being only GHG tables
    print()
    print("Working on Imputing file " + name)
    print()
    # For loop to impute each table separately
    for x in tables1:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 29, name, x)

    for x in tables2:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 23, name, x)

    for x in tables3:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 25, name, x)

    for x in tables4:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 32, name, x)

    for x in tables5:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 23, name, x)
        onedfImp(first_col, 35, last_col, 35, name, x)
        onedfImp(first_col, 36, last_col, 36, name, x)

    for x in tables6:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 21, name, x)
        onedfImp(first_col, 33, last_col, 33, name, x)

    for x in tables7:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 20, name, x)
        onedfImp(first_col, 31, last_col, 31, name, x)

    for x in tables8:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 25, name, x)
        onedfImp(first_col, 39, last_col, 39, name, x)
        onedfImp(first_col, 40, last_col, 40, name, x)

    for x in tables9:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 18, name, x)

    for x in tables10:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)

    for x in tables11:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 16, name, x)

    for x in tables12:
        print("Working on " + x)
        onedfImp(first_col, 13, last_col, 13, name, x)

    for x in tables13:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 17, name, x)
        onedfImp(first_col, 25, last_col, 25, name, x)

    for x in tables14:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 22, name, x)
        onedfImp(first_col, 33, last_col, 33, name, x)
        onedfImp(first_col, 34, last_col, 34, name, x)

    for x in tables15:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 20, name, x)
        onedfImp(first_col, 29, last_col, 29, name, x)
        onedfImp(first_col, 30, last_col, 30, name, x)

    for x in tables16:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 18, name, x)
        onedfImp(first_col, 27, last_col, 27, name, x)

    for x in tables17:
        print("Working on " + x)
        onedfImp(first_col, 13, last_col, 13, name, x)
        onedfImp(first_col, 16, last_col, 16, name, x)
        onedfImp(first_col, 27, last_col, 29, name, x)

    for x in tables18:
        print("Working on " + x)
        onedfImp(first_col, 14, last_col, 14, name, x)
        onedfImp(first_col, 17, last_col, 17, name, x)
        onedfImp(first_col, 20, last_col, 20, name, x)
        onedfImp(first_col, 23, last_col, 23, name, x)
        onedfImp(first_col, 24, last_col, 24, name, x)

    for x in tables19:
        print("Working on " + x)
        onedfImp(first_col, 14, last_col, 14, name, x)
        onedfImp(first_col, 15, last_col, 15, name, x)
        onedfImp(first_col, 16, last_col, 16, name, x)
        onedfImp(first_col, 24, last_col, 24, name, x)
        onedfImp(first_col, 25, last_col, 25, name, x)
        onedfImp(first_col, 26, last_col, 26, name, x)

    for x in tables20:
        print("Working on " + x)
        onedfImp(first_col, 14, last_col, 14, name, x)
        onedfImp(first_col, 15, last_col, 15, name, x)
        onedfImp(first_col, 16, last_col, 16, name, x)
        onedfImp(first_col, 17, last_col, 17, name, x)
        onedfImp(first_col, 26, last_col, 26, name, x)
        onedfImp(first_col, 27, last_col, 27, name, x)
        onedfImp(first_col, 28, last_col, 28, name, x)
        onedfImp(first_col, 29, last_col, 29, name, x)
        onedfImp(first_col, 38, last_col, 38, name, x)
        onedfImp(first_col, 39, last_col, 39, name, x)
        onedfImp(first_col, 40, last_col, 40, name, x)
        onedfImp(first_col, 41, last_col, 41, name, x)
        onedfImp(first_col, 45, last_col, 45, name, x)
        onedfImp(first_col, 46, last_col, 46, name, x)
        onedfImp(first_col, 49, last_col, 49, name, x)
        onedfImp(first_col, 50, last_col, 50, name, x)
        onedfImp(first_col, 53, last_col, 53, name, x)
        onedfImp(first_col, 54, last_col, 54, name, x)
        onedfImp(first_col, 57, last_col, 57, name, x)

for name in filenameCa:
    # define table list which is used in for loop
    tables1 = ["Table 1"]
    tables2 = ["Table 2"]
    tables3 = ["Table 3"]
    tables4 = ["Table 7"]
    tables5 = ["Table 9"]
    tables6 = ["Table 10", "Table 11", "Table 17", "Table 18", "Table 23", "Table 24", "Table 26", "Table 27",
               "Table 29", "Table 31", "Table 35", "Table 41", "Table 42", "Table 44", "Table 46", "Table 48",
               "Table 50", "Table 55", "Table 56", "Table 58", "Table 59"]
    tables7 = ["Table 12"]
    tables8 = ["Table 13", "Table 33", "Table 45"]
    tables9 = ["Table 14", "Table 30", "Table 37", "Table 38", "Table 43", "Table 52", "Table 53"]
    tables10 = ["Table 15"]
    tables11 = ["Table 16", "Table 22", "Table 25", "Table 40", "Table 54"]
    tables12 = ["Table 19"]
    tables13 = ["Table 20", "Table 21", "Table 28"]
    tables14 = ["Table 34"]
    tables15 = ["Table 36", "Table 51"]
    tables16 = ["Table 39"]
    tables17 = ["Table 47", "Table 57"]
    tables18 = ["Transportation1"]
    tables19 = ["Passenger1"]
    tables20 = ["Passenger3"]
    tables21 = ["Freight1"]
    tables22 = ["Freight3"]
    # Ignore table 4, 5, 6, 8, Transportation 2, Passenger 2, Freight 2 due to being only GHG tables
    # Ignore table 32, 49, 60, Passenger 4, Freight 4, Transportation 3
    print()
    print("Working on Imputing file " + name)
    print()
    # For loop to impute each table separately
    for x in tables1:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 29, name, x)
        onedfImp(first_col, 46, last_col, 46, name, x)
        onedfImp(first_col, 47, last_col, 47, name, x)

    for x in tables2:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 23, name, x)
        onedfImp(first_col, 37, last_col, 37, name, x)

    for x in tables3:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 25, name, x)
        onedfImp(first_col, 41, last_col, 41, name, x)

    for x in tables4:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 32, name, x)
        onedfImp(first_col, 52, last_col, 52, name, x)
        onedfImp(first_col, 53, last_col, 53, name, x)

    for x in tables5:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 27, name, x)
        onedfImp(first_col, 42, last_col, 42, name, x)
        onedfImp(first_col, 43, last_col, 43, name, x)

    for x in tables6:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 24, name, x)
        onedfImp(first_col, 39, last_col, 39, name, x)

    for x in tables7:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 23, name, x)
        onedfImp(first_col, 35, last_col, 35, name, x)
        onedfImp(first_col, 36, last_col, 36, name, x)

    for x in tables8:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 21, name, x)
        onedfImp(first_col, 33, last_col, 33, name, x)

    for x in tables9:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 20, name, x)
        onedfImp(first_col, 31, last_col, 31, name, x)

    for x in tables10:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 25, name, x)
        onedfImp(first_col, 39, last_col, 39, name, x)
        onedfImp(first_col, 40, last_col, 40, name, x)

    for x in tables11:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 26, name, x)
        onedfImp(first_col, 41, last_col, 41, name, x)
        onedfImp(first_col, 42, last_col, 42, name, x)

    for x in tables12:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 18, name, x)
        onedfImp(first_col, 25, last_col, 25, name, x)
        onedfImp(first_col, 25, last_col, 26, name, x)

    for x in tables13:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        onedfImp(first_col, 23, last_col, 23, name, x)

    for x in tables14:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 17, name, x)
        onedfImp(first_col, 25, last_col, 25, name, x)

    for x in tables15:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 22, name, x)
        onedfImp(first_col, 33, last_col, 33, name, x)
        onedfImp(first_col, 34, last_col, 34, name, x)

    for x in tables16:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 14, last_col, 15, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 17, last_col, 20, name, x)
        onedfImp(first_col, 29, last_col, 29, name, x)
        onedfImp(first_col, 30, last_col, 30, name, x)

    for x in tables17:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 18, name, x)
        onedfImp(first_col, 27, last_col, 27, name, x)

    for x in tables18:
        print("Working on " + x)
        twodfImp(first_col, 10, last_col, 10, first_col, 11, last_col, 13, name, x)
        twodfImp(first_col, 10, last_col, 10, first_col, 15, last_col, 25, name, x)
        twodfImp(first_col, 10, last_col, 10, first_col, 27, last_col, 41, name, x)
        onedfImp(first_col, 44, last_col, 44, name, x)
        onedfImp(first_col, 45, last_col, 45, name, x)

    for x in tables19:
        print("Working on " + x)
        twodfImp(first_col, 10, last_col, 10, first_col, 12, last_col, 20, name, x)
        twodfImp(first_col, 10, last_col, 10, first_col, 22, last_col, 29, name, x)
        twodfImp(first_col, 32, last_col, 32, first_col, 34, last_col, 41, name, x)

    for x in tables20:
        print("Working on " + x)
        twodfImp(first_col, 10, last_col, 10, first_col, 12, last_col, 18, name, x)
        onedfImp(first_col, 21, last_col, 21, name, x)

    for x in tables21:
        print("Working on " + x)
        twodfImp(first_col, 10, last_col, 10, first_col, 12, last_col, 21, name, x)
        twodfImp(first_col, 10, last_col, 10, first_col, 23, last_col, 28, name, x)
        twodfImp(first_col, 31, last_col, 31, first_col, 33, last_col, 38, name, x)

    for x in tables22:
        print("Working on " + x)
        twodfImp(first_col, 10, last_col, 10, first_col, 12, last_col, 17, name, x)
        onedfImp(first_col, 20, last_col, 20, name, x)

# Define time at end of program
GlobalEndTime = time.time()
# Determine the time and convert to minutes and seconds
TRANTimeMin, TRANTimeSec = divmod((GlobalEndTime - GlobalStartTime) / 60, 1.0)

print()
print(
    "------------------------------------------------------------------------------------------------------------------------")
# Defining the variable for the amount of missing values
TRA_Miss = sum(misslist)
print("There were " + str(TRA_Miss) + " missing values overall")

if end_variable:
    print("Overall completion time: " + str(round(TRANTimeMin)) + " Minutes and " + str(round(TRANTimeSec * 60)) +
          " Seconds")
else:
    print("Transportation Sector completion time: " + str(round(TRANTimeMin)) + " Minutes and " + str(
        round(TRANTimeSec * 60)) + " Seconds")

print("Hopefully completed imputing transportation files successfully")

if end_variable:
    input("Press any key to exit console")
else:
    print()
    print(
        "------------------------------------------------------------------------------------------------------------------------")
