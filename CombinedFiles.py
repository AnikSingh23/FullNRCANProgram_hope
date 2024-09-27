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
    # Get row names in both old and new DataFrames (Column B, index 1)
    old_row_names = old_df.iloc[:, 1].fillna("").tolist()  # Column B is index 1
    new_row_names = new_df.iloc[:, 1].fillna("").tolist()

    for i, new_row_name in enumerate(new_row_names):
        # Skip rows containing "!"
        if "!" in new_row_name:
            print(f"Skipping row with '!': {new_row_name}")
            continue

        if new_row_name not in old_row_names:
            print(f"Inserting missing row: {new_row_name}")

            # Get the new row data
            new_row_data = new_df.iloc[i].tolist()

            # Create a blank row for the old_df if it doesn't have enough columns
            old_row_data = [''] * len(old_df.columns)

            # Ensure the new row has the same number of columns as the old DataFrame
            new_row_data += [0] * (len(old_df.columns) - len(new_row_data))

            # Ensure the old_df row has enough columns as the new DataFrame
            old_row_data += [0] * (len(new_df.columns) - len(old_row_data))

            # Set row name in Column B (index 1)
            new_row_data[1] = new_row_name
            old_row_data[1] = new_row_name

            # Insert the new rows in both DataFrames at the right position
            insert_position = len(old_df)

            for j, old_row_name in enumerate(old_row_names):
                if new_row_name > old_row_name:
                    insert_position = j + 1

            # Insert rows and update the row names
            old_df = pd.concat([old_df.iloc[:insert_position], pd.DataFrame([old_row_data], columns=old_df.columns), old_df.iloc[insert_position:]], ignore_index=True)
            new_df = pd.concat([new_df.iloc[:insert_position], pd.DataFrame([new_row_data], columns=new_df.columns), new_df.iloc[insert_position:]], ignore_index=True)

            old_row_names.insert(insert_position, new_row_name)
            print(f"Row {new_row_name} added at position {insert_position}.")

    return old_df, new_df


# Processing files by province
def process_files_by_province(province_files, province_code):
    for file in province_files:
        print(file)
        LeapName = LeapNameChange(file)
        print(LeapName)
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

                        # Load dataframes from the old and new files
                        old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                        new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                        # Try processing with the current rows
                        process_sheet(old_df, new_df, oldFileName, sheet_name)

                except Exception as e:
                    print(f"An error occurred while processing sheet {sheet_name}: {e}")
                    print(f"Attempting to add missing rows and retry...")

                    # Try adding missing rows to the old DataFrame and retrying
                    try:
                        # Reload old_df in case it was modified before error occurred
                        old_df = pd.read_excel(oldFileName, sheet_name=sheet_name, header=None)
                        new_df = pd.read_excel(file, sheet_name=sheet_name, header=None)

                        # Add missing rows to the old file
                        old_df, new_df = add_missing_rows(old_df, new_df)

                        # Retry processing with added rows
                        process_sheet(old_df, new_df, oldFileName, sheet_name)

                    except Exception as retry_error:
                        print(f"Retry failed for sheet {sheet_name}: {retry_error}")

        except Exception as e:
            print(f"An error occurred with the file {file}: {e}")


def process_sheet(old_df, new_df, oldFileName, sheet_name):
    try:
        # Year row is assumed to be row 9 in both old and new DataFrames (for years starting at index 2)
        old_year_row = old_df.iloc[9, 2:].dropna().astype(int).tolist()
        new_year_row = new_df.iloc[9, 2:].dropna().astype(int).tolist()

        # Find where the years start and how many columns to copy
        old_start_col = old_year_row.index(1990)  # Assuming 1990 is present in both
        num_years_to_copy = len(new_year_row)

        # Adjust the old_df to fit the new columns if needed
        old_required_columns = old_start_col + num_years_to_copy
        if old_df.shape[1] < old_required_columns:
            for _ in range(old_required_columns - old_df.shape[1]):
                old_df[old_df.shape[1]] = 0  # Fill missing columns with 0s

        # Extract the new data (ignoring headers, assuming row 11 onwards is data)
        new_data = new_df.iloc[11:, 2:2 + num_years_to_copy].values

        # Ensure the old_df has enough rows to match new_df, or add missing rows
        current_old_rows = old_df.shape[0]
        required_old_rows = max(old_df.shape[0], new_data.shape[0] + 11)
        if current_old_rows < required_old_rows:
            # Add extra rows if the old file has fewer rows than the new data
            for _ in range(required_old_rows - current_old_rows):
                old_df.loc[current_old_rows] = [''] * old_df.shape[1]
                current_old_rows += 1

        # Now we need to add rows from new_df that are not present in old_df
        old_row_names = old_df.iloc[:, 1].fillna("").tolist()  # Assuming column B contains row names
        new_row_names = new_df.iloc[:, 1].fillna("").tolist()

        for i, new_row_name in enumerate(new_row_names):
            if new_row_name not in old_row_names:
                print(f"Inserting missing row: {new_row_name}")

                # Insert the missing row into old_df
                new_row_data = new_df.iloc[i, :].tolist()
                new_row_data += [0] * (len(old_df.columns) - len(new_row_data))  # Extend row data if needed

                insert_position = len(old_df)
                for j, old_row_name in enumerate(old_row_names):
                    if new_row_name > old_row_name:
                        insert_position = j + 1

                old_df = pd.concat([old_df.iloc[:insert_position], pd.DataFrame([new_row_data], columns=old_df.columns), old_df.iloc[insert_position:]], ignore_index=True)

                old_row_names.insert(insert_position, new_row_name)

        # Insert the new year data and values into the aligned columns of old_df
        old_df.iloc[9, old_start_col:old_start_col + num_years_to_copy] = new_year_row[:num_years_to_copy]
        old_df.iloc[11:11 + new_data.shape[0], old_start_col:old_start_col + num_years_to_copy] = new_data

        # Save the updated DataFrame back to the old file
        with pd.ExcelWriter(oldFileName, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            old_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    except Exception as e:
        raise ValueError(f"An error occurred while processing the sheet: {e}")


# Now process each province's files using the modified code
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
