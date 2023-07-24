import openpyxl
import os

# Path of Excel file
path = r'C:\Users\HaitemElAaouani\Downloads\RawProfiles'
xlsx_file = os.path.join(path, 'GPC_profiles.xlsx')

#Profile to process
uneditedSheet = 'AllergyIntolerance (Unedited)'

# Function to delete reds
def delete_any_red(file_path, unedited_sheet):
    workbook = openpyxl.load_workbook(file_path)
    edited_sheet = workbook.copy_worksheet(workbook[unedited_sheet])
    edited_sheet.title = edited_sheet.title.replace('(Unedited) Copy', '(Edited)')

    for row_idx in range(edited_sheet.max_row + 1, 0, -1):
        if edited_sheet.cell(row=row_idx, column=3).fill.start_color.index == "FFFF0000":
            edited_sheet.delete_rows(row_idx)

    for col_idx in range(edited_sheet.max_column + 1, 0, -1):
        if edited_sheet.cell(row=1, column=col_idx).fill.start_color.index == "FFFF0000":
            edited_sheet.delete_cols(col_idx)

    # Hide Unedited worksheets
    for sheet_name in workbook.sheetnames:
        if sheet_name.endswith('(Unedited)'):
            worksheet = workbook[sheet_name]
            worksheet.sheet_state = "hidden"

    workbook.save(file_path)

# Cleansing of the profile
delete_any_red(xlsx_file, uneditedSheet)