import os
import glob
import sys
import zipfile
import Variables
from Variables import temp_folder


# A variable to define a string to append to the files
YearUnder = Variables.YearInput
# Creates a list of any .zip file in the TEMP folder
ZipFiles = glob.glob(temp_folder + "\*.zip")

for file in ZipFiles:
    zip_file = zipfile.ZipFile(file)  # Defining the zip file variable
    zip_file.extractall(temp_folder)  # Extracting the zip file to temp folder
    UnZip = zip_file.namelist()  # Gets the name of the unzipped file
    TempUnZipPath = temp_folder + "\\" + UnZip[0]  # defines the path to the unzipped file
    SplitUnZip = os.path.splitext(UnZip[0])  # gets the name of the unzipped file without the extension
    TempUnZipPathNoExt = os.path.splitext(TempUnZipPath)  # Defines the path to the unzipped file without the extension
    TempUnZipPathNewName = TempUnZipPathNoExt[0] + YearUnder + ".xls"  # Creates a new path name with included extra
    # text from user input
    os.replace(TempUnZipPath, TempUnZipPathNewName)  # Renames unzipped file to new name with user input appended
    zip_file.close()  # Closes zip file so it can be deleted and etc
    os.remove(file)  # deletes zip file
