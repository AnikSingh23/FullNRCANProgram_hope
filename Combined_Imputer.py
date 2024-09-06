import Variables
from Variables import end_variable
import time

# Start time recording
CompleteStartTime = time.time()
# Change the end_variable for the Variable.py file so every imputer will not ask for an end prompt
Variables.end_variable = False
# Change the end_variable for this specific file since Variables.py was already imported so for intensive purposes
# end_variable is still True for this file. This step is unnecessary as the code at the bottom can be change to
# accommodate it being true, but I prefer keeping the variable the same for all files otherwise it can cause confusion

# Runs the imputers in order
import AGG_Imputer
import AGR_Imputer
import COM_Imputer
import IND_Imputer
import RES_Imputer
import TRA_Imputer

# Imports the variables after the imputer has run to collect information
from AGG_Imputer import AGGTimeMin, AGGTimeSec, AGG_Miss
from AGR_Imputer import AGRTimeMin, AGRTimeSec, AGR_Miss
from COM_Imputer import COMTimeMin, COMTimeSec, COM_Miss
from IND_Imputer import INDTimeMin, INDTimeSec, IND_Miss
from RES_Imputer import RESTimeMin, RESTimeSec, RES_Miss
from TRA_Imputer import TRANTimeMin, TRANTimeSec, TRA_Miss

# Sum all the missing values to give a total value
Total_miss = AGG_Miss + AGR_Miss + COM_Miss + IND_Miss + RES_Miss + TRA_Miss
# Define time at end of program
CompleteEndTime = time.time()
# Determine the time and convert to minutes and seconds
CompleteTimeMin, CompleteTimeSec = divmod((CompleteEndTime - CompleteStartTime) / 60, 1.0)

print()
print("Hopefully completed imputing ALL files successfully")
print()
# AGG info
print(
    "Aggregated industries completion time: " + str(round(AGGTimeMin)) + " Minutes and " + str(round(AGGTimeSec * 60)) +
    " Seconds")
print("Overall amount of missing values for aggregated industries sector: " + str(AGG_Miss) + " Missing Values")
print()
# IND info
print("Disaggregated industries completion time: " + str(round(INDTimeMin)) + " Minutes and " + str(
    round(INDTimeSec * 60)) +
      " Seconds")
print("Overall amount of missing values for disaggregated industries sector: " + str(IND_Miss) + " Missing Values")
print()
# AGR Info
print("Agricultural completion time: " + str(round(AGRTimeMin)) + " Minutes and " + str(round(AGRTimeSec * 60)) +
      " Seconds")
print("Overall amount of missing values for agricultural sector: " + str(AGR_Miss) + " Missing Values")
print()
# COM Info
print("Commercial and institutional completion time: " + str(round(COMTimeMin)) + " Minutes and " + str(
    round(COMTimeSec * 60)) +
      " Seconds")
print("Overall amount of missing values for commercial and institutional sector: " + str(COM_Miss) + " Missing Values")
print()
# RES Info
print("Residential completion time: " + str(round(RESTimeMin)) + " Minutes and " + str(round(RESTimeSec * 60)) +
      " Seconds")
print("Overall amount of missing values for residential sector: " + str(RES_Miss) + " Missing Values")
print()
# TRA Info
print("Transportation completion time: " + str(round(TRANTimeMin)) + " Minutes and " + str(round(TRANTimeSec * 60)) +
      " Seconds")
print("Overall amount of missing values for transportation sector: " + str(TRA_Miss) + " Missing Values")
print()
# Overall Info
print("Overall Completion time of every file: " + str(round(CompleteTimeMin)) + " Minutes and " + str(
    round(CompleteTimeSec * 60)) + " Seconds")
print("Overall missing values of every file: " + str(Total_miss) + " Missing Values")
print()
print()
