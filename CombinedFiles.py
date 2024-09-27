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

# Dynamic function to add missing rows and columns
def dynamically_add_missing_rows_columns(old_df, new_df):
    # Find how many rows and columns need to be added in the old_df to match new_df
    old_row_count, old_col_count = old_df.shape
    new_row_count, new_col_count = new_df.shape

    # Dynamically add missing rows if old_df has fewer rows than new_df
    if old_row_count < new_row_count:
        # Add empty rows to old_df to match new_df row count
        rows_to_add = new_row_count - old_row_count
        new_rows = pd.DataFrame([[''] * old_col_count] * rows_to_add)  # Create empty rows
        old_df = pd.concat([old_df, new_rows], ignore_index=True)
        print(f"Added {rows_to_add} rows to match new file.")

    # Dynamically add missing columns if old_df has fewer columns than new_df
    if old_col_count < new_col_count:
        # Add empty columns to old_df to match new_df column count
        cols_to_add = new_col_count - old_col_count
        for i in range(cols_to_add):
            old_df[f'new_col_{i+1}'] = ''  # Dynamically name the new columns
        print(f"Added {cols_to_add} columns to match new file.")

    return old_df

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
    LeapName = LeapNameChange(file)
    print(LeapName)
    oldFileName = source_folder + "\\AB\\" + LeapName
    newFileName = source_folder + "\\AB\\" + LeapName[:-5] + " " + str(OldYearInput) + ".xlsx"

    try:
        # Load the entire workbook for both old and new files
        old_workbook = pd.ExcelFile(oldFileName)
        new_workbook = pd.ExcelFile(file)

        # Iterate through the sheet names in the old workbook
        for sheet_name in old_workbook.sheet_names:
            try:
                if pattern.match(sheet_name):
                    print(f"Processing {sheet_name}...")

                    # Load the specific sheet (table) from both the old and new files
                    old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                    new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                    # Dynamically add missing rows/columns
                    old_df = dynamically_add_missing_rows_columns(old_df, new_df)

                    # Extract the year row (assumed to be in row 9 in the old file and row 10 in the new file)
                    old_year_row = old_df.iloc[9, 2:]
                    new_year_row = new_df.iloc[10, 2:]

                    # Ensure both year rows have the expected structure
                    old_years = old_year_row.dropna().astype(int).tolist()
                    new_years = new_year_row.dropna().astype(int).tolist()

                    # Determine the starting column for year 2000 in the old file
                    old_start_col = old_years.index(2000) + 2
                    num_years_to_copy = len(new_years)

                    old_required_columns = old_start_col + num_years_to_copy
                    current_old_columns = len(old_df.columns)

                    # Add extra columns if the old file doesn't have enough columns (repeated in case dynamic adjustment missed something)
                    if current_old_columns < old_required_columns:
                        for _ in range(old_required_columns - current_old_columns):
                            old_df[current_old_columns] = ''
                            current_old_columns += 1

                    # Debug info
                    print(f"Old Start Col: {old_start_col}, Columns to Copy: {num_years_to_copy}, Required Columns: {old_required_columns}, Current Columns: {current_old_columns}")

                    # Update the year row (row 9) in the old file to include missing years from 2000-2021
                    old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row.values[:num_years_to_copy]

                    # Copy data from new file to old file
                    new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values
                    old_df.iloc[10:10 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

                    # Save the modified sheet back to the old workbook, preserving other sheets
                    with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            except Exception as e:
                print(f"An error occurred while processing sheet {sheet_name}: {e}")

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
