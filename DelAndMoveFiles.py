import Variables
import glob
import shutil
import os
import pandas as pd
from Variables import temp_folder, source_folder, YearInput, LeapNameChange, OldYearInput, Province

OldFiles = glob.glob(temp_folder + "\*.xls")
for file in OldFiles:
    os.remove(file)

AB_File = glob.glob(temp_folder + "\*ab*.xlsx")

ATL_File = glob.glob(temp_folder + "\*atl*.xlsx")

CAN_File = glob.glob(temp_folder + "\*ca*.xlsx")

BC_File = glob.glob(temp_folder + "\*BC*.xlsx")

MB_File = glob.glob(temp_folder + "\*MB*.xlsx")

ON_File = glob.glob(temp_folder + "\*ON*.xlsx")

QC_File = glob.glob(temp_folder + "\*QC*.xlsx")

SK_File = glob.glob(temp_folder + "\*SK*.xlsx")

NB_File = glob.glob(temp_folder + "\*NB*.xlsx")

NL_File = glob.glob(temp_folder + "\*NF*.xlsx")

PE_File = glob.glob(temp_folder + "\*PE*.xlsx")

NS_File = glob.glob(temp_folder + "\*NS*.xlsx")

New_File = glob.glob(temp_folder + "\*.xlsx")

# Remove the temp folder from the province list
Province.remove("Temp")

# Define the input for the latest year of the old files
# if OldYearInput == "*":
#     # Set a variable for the while loop to have an exit in case there are no old files to replace
#     i = 0
#     # While loop with the exit conditions of the the OldYearInput becoming literally anything else or has been processed 50 times
#     while OldYearInput == "*" or i > 50:
#         # Define a variable for later use
#         newFileName = ""
#         # Using try here just incase the file does not exist because it would throw an error if it does not, basically
#         # what this does is take a file from the temp folder change it name to match with leap, determine if it already
#         # exists, and extract the year it was updated until and use that year for the rest of the name changes.
#         # Also, the original files are deleted by this point so there is no need to worry about the name formatting of
#         # those files
#         try:
#             LeapName = LeapNameChange(New_File[i])  # Creates the Leap Name
#             # Iterate through the province folders until a match is found
#             print(LeapName)
#             for p in Province:
#                 try:
#                     newFileName = source_folder + "\\" + p + "\\" + LeapName  # Finds the files by searching the province folders
#                     if os.path.exists(newFileName):  # If filename is not blank it will break out of the loop
#                         break
#                 except (Exception,):
#                     pass  # If an exception occurs it will automatically pass and restart the for loop
#
#             print(newFileName)
#             dfcheck1 = pd.read_excel(newFileName, sheet_name="Table 1", skiprows=10, nrows=0)  # Read the rows with the years if the format is the same
#             dfcheck2 = pd.read_excel(newFileName, sheet_name="Table 1", skiprows=9, nrows=0)  # Read the rows with the year data from previous format
#
#             if max(dfcheck1, dfcheck2) > 3000:  # if the higher number from the dfchecks is greater than 3000 use the
#                 # lower value. Generally the higher value of the two will be the updated year but just incase there is an
#                 # outlying number this will prevent some issues
#                 dfcheck = min(dfcheck1, dfcheck2)
#             else:
#                 dfcheck = max(dfcheck1, dfcheck2)
#             print(dfcheck)
#             year_list = dfcheck.columns.tolist()  # Turn the single column into a list
#             OldYearInput = str(year_list[
#                                    -1])  # Grab the last entry in the list which should always be the latest year and convert it into a string
#         except (Exception,):  # If an exception occurs (when a file does not exist) increment the i variable and continue the while loop
#             break
#     i = i + 1
#     if OldYearInput == "*":  # If oldyearinput is not defined by the end of the while loop (50 iterations) it will make the variable the string "undefined"
#         OldYearInput = "Undefined"
if OldYearInput == "*":
    for i in New_File:
        LeapName = LeapNameChange(i)

        for p in Province:
            newFileName = source_folder + "\\" + p + "\\" + LeapName
            if os.path.exists(newFileName):
                break
        else:
            continue
        try:
            dfcheck1 = pd.read_excel(newFileName, sheet_name="Table 1", skiprows=10, nrows=0)
            dfcheck2 = pd.read_excel(newFileName, sheet_name="Table 1", skiprows=9, nrows=0)
        except (Exception,):
            pass
        dfcheck1 = dfcheck1.columns.tolist()
        dfcheck1 = dfcheck1[2:]
        if not dfcheck1:
            dfcheck1.append(1)
        dfcheck2 = dfcheck2.columns.tolist()
        dfcheck2 = dfcheck2[2:]
        if not dfcheck2:
            dfcheck2.append(1)
        if max(dfcheck1) == 1 and max(dfcheck2) == 1:
            break
        if max(max(dfcheck1), max(dfcheck2)) > 3000:
            dfcheck = min(max(dfcheck1), max(dfcheck2))
        else:
            dfcheck = max(max(dfcheck1), max(dfcheck2))
        OldYearInput = str(dfcheck)
        if OldYearInput != "Undefined" or OldYearInput != 1:
            break


for file in AB_File:
    LeapName = LeapNameChange(file)  # Use a function defined in variables.py to change the file name to a leap compliant version
    newFileName = source_folder + "\\AB\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\AB\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"  # The [:-5] removes the .xlsx

    try:  # try to rename an old file if possible and if the file name already exists then increment a number and append that to the filename
        i = 1  # Set the increment variable which will be used incase the name is already in use (starts from 1 rather than 0)
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\AB\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in ATL_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\ATL\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\ATL\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\ATL\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in CAN_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\CAN\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\CAN\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\CAN\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in BC_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\BC\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\BC\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\BC\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in MB_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\MB\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\MB\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\MB\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in NB_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\NB\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\NB\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\NB\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in NL_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\NL\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\NL\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\NL\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in NS_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\NS\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\NS\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\NS\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in ON_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\ON\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\ON\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\ON\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in PE_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\PE\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\PE\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\PE\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in QC_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\QC\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\QC\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\QC\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

for file in SK_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\SK\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\SK\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\SK\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate

TER_File = glob.glob(temp_folder + "\*TR*.xlsx")

for file in TER_File:
    LeapName = LeapNameChange(file)
    newFileName = source_folder + "\\TER\\" + LeapName  # Creates file path for new folder
    oldFileName = source_folder + "\\TER\\" + LeapName[:-5] + " " + OldYearInput + ".xlsx"
    try:
        i = 1
        while os.path.exists(oldFileName):
            oldFileName = source_folder + "\\TER\\" + LeapName[:-5] + " " + OldYearInput + " " + "(" + str(
                i) + ")" + ".xlsx"
            i = i + 1
        shutil.move(newFileName, oldFileName)  # Moves old and renames old files
    except (Exception,):
        pass
    shutil.move(file, newFileName)  # Moves newly unzipped and renamed file to new path in appropriate
