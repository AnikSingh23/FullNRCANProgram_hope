import time  # Generally you would import everything at the same time, but I want the program to run in a specific order to I am ordering the imports as necessary

StartTime = time.time()  # Start the overall timer
print("Starting program which will download, unzip, impute, and move files into folders consistent with LEAP")
print()

import Variables  # Run the Variables script to get the functions, extra input and etc
import NRCANScraperDownloader  # Run the downloader to download files to a temporary folder

print()
print("Downloaded Files will proceed to unzip")
import Unzip_Files  # Unzip the files in the temporary folder

print()
print("Unzipped Files will proceed to convert and impute the files")
print()
import Column_Check  # Check the first and last year columns which will then be used in each of the imputer files
from Column_Check import first_year, last_year  # Import the variables to print out which years the files are dated

print()
print("These excel files are from the years " + str(first_year) + " to " + str(last_year))
print()
import Combined_Imputer  # Run the Combined Imputer script which will run each imputer in order and determine the amount of missing values and time each section took

print()
print("Imputation Completed will proceed to del old .xls files and move the .xlsx files")
# import DelAndMoveFiles  # Run the py file which will delete the old files and move the new files
import CombinedFiles # Run to combine files into a single large file.
print()
print("Finished moving and deleting files")
print()
# Timing function
EndTime = time.time()  # Record time at end of script
# Determine the time and convert to minutes and seconds
CompleteTimeMin, CompleteTimeSec = divmod((EndTime - StartTime) / 60, 1.0)
print("The Whole process took: " + str(round(CompleteTimeMin)) + " Minutes and " + str(
    round(CompleteTimeSec * 60)) + " Seconds")
print()
print()
input("Press enter to exit console/program")  # Gives an input so the console just does not exit abruptly
