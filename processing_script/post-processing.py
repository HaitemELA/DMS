import openpyxl
from openpyxl.styles import PatternFill
import os

# Path of Excel file
path = r'C:\Users\HaitemElAaouani\Documents\DMS\RawProfiles'
xlsx_file = os.path.join(path, 'GPC_profiles.xlsx')

#Profile to process
uneditedSheet = 'Medication (Unedited)'

# Cornflower Blue	 
blue_fill = PatternFill(start_color='FF6D9EEB', #4A86E8
                           end_color='FF6D9EEB',
                           fill_type='solid')

# WorkBook definition
wkbk = openpyxl.load_workbook(xlsx_file)

# Function to delete reds
def delete_any_red(workbook, unedited_sheet):
    edited_sheet = workbook.copy_worksheet(workbook[unedited_sheet])
    edited_sheet.title = edited_sheet.title.replace('(Unedited) Copy', '(Edited)')

    for row_idx in range(edited_sheet.max_row + 1, 0, -1):
        if edited_sheet.cell(row=row_idx, column=3).fill.start_color.index == "FFFF0000":
            edited_sheet.delete_rows(row_idx)

    for col_idx in range(edited_sheet.max_column + 1, 0, -1):
        if edited_sheet.cell(row=1, column=col_idx).fill.start_color.index == "FFFF0000":
            edited_sheet.delete_cols(col_idx)
        else:
            edited_sheet.cell(row=1, column=col_idx).fill = blue_fill

    return edited_sheet.title

#Copy range of cells as a nested list
#Takes: start cell, end cell, and sheet you want to copy from.
def ElementName_FhirTarget(workbook, Edited_sheet, source_col, destination_col):
    sheet = workbook[Edited_sheet]
    #Loops through selected Rows
    for i in range(2,sheet.max_row + 1):
        sheet.cell(row = i, column = destination_col).value = sheet.cell(row = i, column = source_col).value
        s = str(sheet.cell(row = i, column = source_col).value)
        s = s.replace('.',' ')
        s = s.replace('extension:','')
        sheet.cell(row = i, column = source_col).value = s
# Hide Unedited worksheets
def hide_unideted(file_path, workbook):
    for sheet_name in workbook.sheetnames:
        if sheet_name.endswith('(Unedited)'):
            worksheet = workbook[sheet_name]
            worksheet.sheet_state = "hidden"

    workbook.save(xlsx_file)

# Cleansing of the profile
EditSht = delete_any_red(wkbk, uneditedSheet)
wkbk[EditSht].cell(row=1, column=1).value = 'Old ID'
wkbk[EditSht].cell(row=1, column=1).fill = blue_fill
wkbk[EditSht].insert_cols(2, amount=1)
wkbk[EditSht].cell(row=1, column=2).value = 'New ID'
wkbk[EditSht].cell(row=1, column=2).fill = blue_fill
wkbk[EditSht].cell(row=1, column=3).value = 'Element Name'
wkbk[EditSht].cell(row=1, column=3).fill = blue_fill
wkbk[EditSht].cell(row=1, column=5).value = 'Description'
wkbk[EditSht].cell(row=1, column=8).value = 'Value Sets'
wkbk[EditSht].cell(row=1, column=8).fill = blue_fill
wkbk[EditSht].cell(row=3, column=8).value = 'IFERROR(index(CIS!$1:$1844,match(I3,CIS!$J$1:$J$1844,0),9),index(CIS!$1:$1844,match(I3,CIS!$J$1:$J$1844,0),9))'
wkbk[EditSht].cell(row=1, column=9).value = 'FHIR Target STU3'
wkbk[EditSht].cell(row=1, column=9).fill = blue_fill
ElementName_FhirTarget(wkbk, EditSht,3,9)
hide_unideted(xlsx_file, wkbk)
