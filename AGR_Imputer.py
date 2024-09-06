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
TAGR_File = []
TAGR_Ca_File = []

# search for files with agr and .xls in it # If a different format is required these searches can be modified
AGR_File = glob.glob(temp_folder + "\\agr*.xls")
# search for files with imp in it (so it doesn't try to impute and already imputed file
ExistingAgrImp = glob.glob(temp_folder + "\\*imp*")
# Add the default canada agr file to add to exclusions
ExistingAgrImp.append(temp_folder + "\\agr_ca_e.xls")
# Add any files with ca and .xls to exclusion list
ExistingAgrImp.extend(glob.glob(temp_folder + "\\*ca*.xls"))

# Removes already imputed files and the canada wide files from the list (first "if" removes any files from exclusion
# list ExistingAgrImp, the for loop will add any files from the AGR_File list which has a ca and .xls in it to its
# own list for further use)
for files in AGR_File:
    if files not in ExistingAgrImp:
        TAGR_File.append(files)
    for s in glob.glob(temp_folder + "\\*ca*"):
        if files == s:
            TAGR_Ca_File.append(files)

# This for loop goes through the filtered files list (this one specifically the files with agr and .xls in somewhere
# and removes the CA file and any file with imp in the file name) note this will create/overwrite any imp files by
# creating them (since imp was excluded the for loop creates a new imp path and adds it to a list to be processed
# later. The reason it adds it to a list no matter what is incase if overwriting is not selected the files will be
# checked anyway (I could create a user input to prevent this, but I don't think that is necessary)
for t in TAGR_File:
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
    elif os.path.splitext(t)[0] + "_imp.xlsx" in ExistingAgrImp:
        print(file_name_imp + " already exists no need to convert")
        filename.append(file_name_imp)
    # This final else catches all files which do not satisfy the previous conditions
    else:
        print("working on converting file " + file_name + " to .xlsx format")
        conversion(file_name, file_name_imp)
        filename.append(file_name_imp)
        print("Finished creating and converting " + file_name_imp)
print()
for t in TAGR_Ca_File:
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
    elif os.path.splitext(t)[0] + "_imp.xlsx" in ExistingAgrImp:
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
    print("Check file, this program will apply the same imputation as the AGR_CA_E.xls which is the only agr file to "
          "exist in the version of the files this was programmed on, if an exception occurs due to differences in "
          "tables the program may need to be slightly altered") 
    tables1 = ["Table 1", "Table 5", "Table 6", "Table 7", "Table 8", "Table 9", "Table 10", "Table 11"]
    tables2 = ["Table 2"]
    tables3 = ["Table 3"]
    tables4 = ["Table 4"]
    print()
    print("Working on Imputing file " + name)
    print()
    # For loop to impute each table separately
    for x in tables1:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 26, name, x)
        onedfImp(first_col, 40, last_col, 40, name, x)
    for x in tables2:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 26, name, x)
    for x in tables3:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 28, name, x)
        onedfImp(first_col, 44, last_col, 44, name, x)
    for x in tables4:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 28, name, x)

for name in filenameCa:

    # create the table lists for the tables which follow the same structure to be used in the following for commands.
    tables1 = ["Table 1", "Table 5", "Table 6", "Table 7", "Table 8", "Table 9", "Table 10", "Table 11"]
    tables2 = ["Table 2"]
    tables3 = ["Table 3"]
    tables4 = ["Table 4"]
    print()
    print("Working on Imputing file " + name)
    print()

    for x in tables1:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 26, name, x)
        onedfImp(first_col, 40, last_col, 40, name, x)
    for x in tables2:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 26, name, x)
    for x in tables3:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 28, name, x)
        onedfImp(first_col, 44, last_col, 44, name, x)
    for x in tables4:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 16, name, x)
        twodfImp(first_col, 13, last_col, 13, first_col, 18, last_col, 28, name, x)
# Define time at end of program
GlobalEndTime = time.time()
# Determine the time and convert to minutes and seconds
AGRTimeMin, AGRTimeSec = divmod((GlobalEndTime - GlobalStartTime) / 60, 1.0)

print()
print(
    "------------------------------------------------------------------------------------------------------------------------")
# Defining the variable for the amount of missing values
AGR_Miss = sum(misslist)
print("There were " + str(AGR_Miss) + " missing values overall")

if end_variable:
    print("Overall completion time: " + str(round(AGRTimeMin)) + " Minutes and " + str(round(AGRTimeSec * 60)) +
          " Seconds")
else:
    print("Agricultural sector completion time: " + str(round(AGRTimeMin)) + " Minutes and " + str(
        round(AGRTimeSec * 60)) + " Seconds")

print("Hopefully completed imputing agricultural files successfully")
if end_variable:
    input("Press any key to exit console")
else:
    print()
    print(
        "------------------------------------------------------------------------------------------------------------------------")
