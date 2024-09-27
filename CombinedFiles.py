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

for file in AB_File:
    print(file)
    LeapName = LeapNameChange(file)  # Use a function defined in variables.py to change the file name to a leap-compliant version
    print(LeapName)
    oldFileName = source_folder + "\\AB\\" + LeapName  # Path for the old file
    newFileName = source_folder + "\\AB\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"  # Path for the new file

    try:
        # Load the entire workbook for both old and new files
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)
        test = []
        # Iterate through the sheet names in the old workbook
        for sheet_name in old_workbook.sheet_names:
            try:
                # Check if the sheet name matches the "Table #" or "Table ##" format
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    # Load the specific sheet (table) from both the old and new files
                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    # Extract the year row (assumed to be in row 9 in the old file and row 10 in the new file)
                    old_year_row = old_df.iloc[9, 2:]  # Years from column C onwards in the old file
                    new_year_row = new_df.iloc[10, 2:]  # Years from column C onwards in the new file

                    # Ensure both year rows have the expected structure (integers or strings representing years)
                    old_years = old_year_row.dropna().astype(int).tolist()  # List of years in the old file
                    new_years = new_year_row.dropna().astype(int).tolist()  # List of years in the new file

                    # Determine the starting column for year 2000 in the old file
                    old_start_col = old_years.index(2000) + 2  # Column index in the old file for year 2000 (offset by 2 for columns starting at C)

                    # Determine the number of years to copy (2000-2021) from the new file
                    num_years_to_copy = len(new_years)  # Copy all columns from 2000-2021 in the new file

                    # Ensure the old file has enough columns for all years from 2000-2021
                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    # Add extra columns if the old file doesn't have enough columns
                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''  # Add empty columns
                            current_old_columns += 1

                    # Debug: print column indices and number of columns to copy
                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    # Update the year row (row 9) in the old file to include missing years from 2000-2021
                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    # Copy data from new file (years 2000-2021, C:X) to old file (M:AH)
                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values  # Shift data up by one row
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data  # Insert into old file

                    # Save the modified sheet back to the old workbook, preserving other sheets
                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")  # Skip to the next sheet if an error occurs

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing ATL_File
for file in ATL_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\ATL\\" + LeapName
    newFileName = source_folder + "\\ATL\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing CAN_File
for file in CAN_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\CAN\\" + LeapName
    newFileName = source_folder + "\\CAN\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing BC_File
for file in BC_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\BC\\" + LeapName
    newFileName = source_folder + "\\BC\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing MB_File
for file in MB_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\MB\\" + LeapName
    newFileName = source_folder + "\\MB\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing NB_File
for file in NB_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\NB\\" + LeapName
    newFileName = source_folder + "\\NB\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")
# Processing NL_File
for file in NL_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\NL\\" + LeapName
    newFileName = source_folder + "\\NL\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing NS_File
for file in NS_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\NS\\" + LeapName
    newFileName = source_folder + "\\NS\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing ON_File
for file in ON_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\ON\\" + LeapName
    newFileName = source_folder + "\\ON\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing PE_File
for file in PE_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\PE\\" + LeapName
    newFileName = source_folder + "\\PE\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing QC_File
for file in QC_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\QC\\" + LeapName
    newFileName = source_folder + "\\QC\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing SK_File
for file in SK_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\SK\\" + LeapName
    newFileName = source_folder + "\\SK\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")

# Processing TER_File
for file in TER_File:
    print(file)
    LeapName = LeapNameChange(file)
    oldFileName = source_folder + "\\TER\\" + LeapName
    newFileName = source_folder + "\\TER\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"An error occurred with the file {file}: {e}")
