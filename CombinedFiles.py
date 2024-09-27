import Variables
import glob
import shutil
import os
import re
import pandas as pd
import openpyxl
from Variables import temp_folder, source_folder, YearInput, LeapNameChange, OldYearInput, Province, checkvalues

OldFiles = glob.glob(temp_folder + "\*.xls")

pattern = re.compile(r'^Table \d{1,2}$')


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

TER_File = glob.glob(temp_folder + "\*TER*.xlsx")


New_File = glob.glob(temp_folder + "\*.xlsx")


# Remove the temp folder from the province list
Province.remove("Temp")

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

def add_missing_rows(old_df, new_df):
    # Find all row names in both old and new DataFrames (row names are in the second column, index 1)
    old_row_names = old_df.iloc[:, 1].fillna("").tolist()  # Column B is index 1
    new_row_names = new_df.iloc[:, 1].fillna("").tolist()

    # Iterate through the new row names and check if they exist in the old row names
    for i, new_row_name in enumerate(new_row_names):
        if new_row_name not in old_row_names:
            print(f"Inserting missing row: {new_row_name}")

            # Get the new row data and fill missing columns with zeros if needed
            new_row_data = new_df.iloc[i].tolist()

            # Ensure the new row has the same number of columns as the old DataFrame
            if len(new_row_data) < len(old_df.columns):
                new_row_data.extend([0] * (len(old_df.columns) - len(new_row_data)))  # Pad with zeros

            # Create a new DataFrame for the row with the correct number of columns
            new_row = pd.DataFrame([new_row_data], columns=old_df.columns)

            # Find where to insert the row in old_df (insert at the same index `i` or at the end)
            if i < len(old_df):
                old_df = pd.concat([old_df.iloc[:i], new_row, old_df.iloc[i:]], ignore_index=True)
            else:
                old_df = pd.concat([old_df, new_row], ignore_index=True)

            print(f"Row {new_row_name} added to old_df.")

    return old_df

# Processing files by province
def process_files_by_province(province_files, province_code):
    for file in province_files:
        print(file)
        LeapName = LeapNameChange(file)
        oldFileName = source_folder + f"\\{province_code}\\" + LeapName
        newFileName = source_folder + f"\\{province_code}\\" + LeapName[:-5] + f" {OldYearInput}.xlsx"

        try:
            # Load the entire workbook for both old and new files
            old_workbook = pd.ExcelFile(oldFileName)
            new_workbook = pd.ExcelFile(file)

            for sheet_name in old_workbook.sheet_names:
                try:
                    if pattern.match(sheet_name):
                        print(f"Processing {sheet_name}...")

                        old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                        new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                        # Add missing rows to the old file
                        old_df = add_missing_rows(old_df, new_df)

                        # Now the rows should match, proceed with copying the year data as needed
                        old_year_row = old_df.iloc[9, 2:]  # Year row assumed at row 10 (index 9)
                        new_year_row = new_df.iloc[10, 2:]

                        old_years = old_year_row.dropna().astype(int).tolist()
                        new_years = new_year_row.dropna().astype(int).tolist()

                        old_start_col = old_years.index(2000) + 2
                        num_years_to_copy = len(new_years)

                        old_required_columns = old_start_col + num_years_to_copy
                        current_old_columns = len(old_df.columns)

                        # Add extra columns if the old file doesn't have enough columns
                        if current_old_columns < old_required_columns:
                            for _ in range(old_required_columns - current_old_columns):
                                old_df[current_old_columns] = ''
                                current_old_columns += 1

                        print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                        # Update the year row (row 9) in the old file to include missing years from 2000-2021
                        old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                        # Copy data from new file (years 2000-2021) to old file
                        new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                        old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                        # Save the modified sheet back to the old workbook, preserving other sheets
                        with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

                except Exception as e:
                    print(f"An error occurred while processing sheet {sheet_name}: {e}")

        except Exception as e:
            print(f"An error occurred with the file {file}: {e}")

# Now process each province's files using the modified code
process_files_by_province(AB_File, "AB")
process_files_by_province(AB_File, "AB")
process_files_by_province(ATL_File, "ATL")
process_files_by_province(CAN_File, "CAN")
process_files_by_province(BC_File, "BC")
process_files_by_province(MB_File, "MB")
process_files_by_province(ON_File, "ON")
process_files_by_province(QC_File, "QC")
process_files_by_province(SK_File, "SK")
process_files_by_province(NB_File, "NB")
process_files_by_province(NL_File, "NL")
process_files_by_province(PE_File, "PE")
process_files_by_province(NS_File, "NS")
process_files_by_province(TER_File, "TER")