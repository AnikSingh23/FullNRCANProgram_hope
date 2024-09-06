import Variables  # Run the Variables script to get the functions, extra input and etc

print()
print("Unzipped Files will proceed to convert and impute the files")
print()
import Column_Check  # Check the first and last year columns which will then be used in each of the imputer files
from Column_Check import first_year, last_year  # Import the variables to print out which years the files are dated

print()
print("These excel files are from the years " + str(first_year) + " to " + str(last_year))
print()
import Combined_Imputer  # Run the Combined Imputer script which will run each imputer in order and determine the amount of missing values and time each section took
