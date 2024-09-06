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

# defining lists to be used for sorting
filename = []
filenameCa = []
TAGG_File = []
TAGG_Ca_File = []

# search for files with agg and .xls in it # If a different format is required these searches can be modified
AGG_File = glob.glob(temp_folder + "\\agg*.xls")
# search for files with imp in it (so it doesn't try to impute and already imputed file
ExistingAggImp = glob.glob(temp_folder + "\\*imp*")
# Add the default canada agg file to add to exclusions
ExistingAggImp.append(temp_folder + "\\agg_ca_e.xls")
# Add any files with ca and .xls to exclusion list
ExistingAggImp.extend(glob.glob(temp_folder + "\\*ca*.xls"))

# Removes already imputed files and the canada wide files from the list (first "if" removes any files from exclusion
# list ExistingAggImp, the for loop will add any files from the AGG_File list which has a ca and .xls in it to its
# own list for further use)
for files in AGG_File:
    if files not in ExistingAggImp:
        TAGG_File.append(files)
    for s in glob.glob(temp_folder + "\\*ca*"):
        if files == s:
            TAGG_Ca_File.append(files)

# This for loop goes through the filtered files list (this one specifically the files with agg and .xls in somewhere
# and removes the CA file and any file with imp in the file name) note this will create/overwrite any imp files by
# creating them (since imp was excluded the for loop creates a new imp path and adds it to a list to be processed
# later. The reason it adds it to a list no matter what is incase if overwriting is not selected the files will be
# checked anyway (I could create a user input to prevent this, but I don't think that is necessary)
for t in TAGG_File:
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
    # This else if statement checks if the file already exists as an imputed.xlsx file and just adds it to the list
    elif os.path.splitext(t)[0] + "_imp.xlsx" in ExistingAggImp:
        print(file_name_imp + " already exists no need to convert")
        filename.append(file_name_imp)
    # This final else catches all files which do not satisfy the previous conditions and converts the process
    else:
        print("working on converting file " + file_name + " to .xlsx format")
        conversion(file_name, file_name_imp)
        filename.append(file_name_imp)
        print("Finished creating and converting " + file_name_imp)

print()

for t in TAGG_Ca_File:
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
    elif os.path.splitext(t)[0] + "_imp.xlsx" in ExistingAggImp:
        print(file_name_imp + " already exists no need to convert")
        filenameCa.append(file_name_imp)
    # This final else catches all files which do not satisfy the previous conditions
    else:
        print("working on converting file " + file_name + " to .xlsx format")
        conversion(file_name, file_name_imp)
        filenameCa.append(file_name_imp)
        print("Finished creating and converting " + file_name_imp)

# Define time at end of conversion
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
    tables = ["Table 1", "Table 2", "Table 3", "Table 4", "Table 5", "Table 6", "Table 7", "Table 8", "Table 9",
              "Table 10", "Table 11", "Table 12"]
    print()
    print("Working on Imputing file " + name)
    print()
    # For loop to impute each table separately
    for x in tables:
        # Calls two data frame imputation function from variables.py the structure of the function takes a range like
        # C13:V13 which is the totals row and C15:V24 and requires the parameters to be broken up into the format as
        # seen (col1,row1,col2,row2,col3,row3,col4,row4,name of file, name of table). Since name of file and sheet
        # were taken care of earlier through the variables 'name' and 'x' from the for loops the only user defined
        # variable is splitting up the C13:V13 and C15:V24 ranges into first_col, 13, last_col, 13, first_col, 15, last_col, 24 sets
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 24, name, x)

for name in filenameCa:

    # create the table lists for the tables which follow the same structure to be used in the following for commands.
    tables1 = ["Table 1", "Table 2", "Table 3", "Table 4", "Table 7", "Table 8", "Table 9", "Table 10", "Table 11",
               "Table 12", "Table 13", "Table 14", "Table 15", "Table 16"]
    tables2 = ["Table 5", "Table 6"]
    print()
    print("Working on Imputing file " + name)
    print()
    for x in tables1:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 24, name, x)
    for x in tables2:
        print("Working on " + x)
        twodfImp(first_col, 13, last_col, 13, first_col, 15, last_col, 21, name, x)

# Define time at end of program
GlobalEndTime = time.time()
# Determine the time and convert to minutes and seconds
AGGTimeMin, AGGTimeSec = divmod((GlobalEndTime - GlobalStartTime) / 60, 1.0)

print()
print(
    "------------------------------------------------------------------------------------------------------------------------")
# Defining the variable for the amount of missing values
AGG_Miss = sum(misslist)
print("There were " + str(AGG_Miss) + " missing values overall")

# This if changes what is printed depending on if this is being run solo or run as part of the combined imputer
if end_variable:
    print("Overall completion time: " + str(round(AGGTimeMin)) + " Minutes and " + str(round(AGGTimeSec * 60)) +
          " Seconds")
else:
    print("Aggregated industries sector completion time: " + str(round(AGGTimeMin)) + " Minutes and " + str(
        round(AGGTimeSec * 60)) + " Seconds")

print("Hopefully completed imputing aggregated industries files successfully")

if end_variable:
    input("Press any key to exit console")
else:
    print()
    print(
        "------------------------------------------------------------------------------------------------------------------------")
