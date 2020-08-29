import xlrd
import xlsxwriter

# Create a workbook and add a worksheet.
primary_wb = xlsxwriter.Workbook('primaryReds.xlsx')
primary_ws = primary_wb.add_worksheet()

secondary_wb = xlsxwriter.Workbook('primaryReds.xlsx')
secondary_ws = primary_wb.add_worksheet()

# Give the location of the file
loc_reds = r"C:\Users\arhon\Google Drive\Recruitment 2020\More_Than_One_Apart.xlsx"
reds_list = xlrd.open_workbook(loc_reds)
red_sheet = reds_list.sheet_by_index(0)

loc_secondary = r"C:\Users\arhon\Google Drive\Recruitment 2020\secondary_list.xlsx"
secondary_list = xlrd.open_workbook(loc_secondary)
secondary_sheet = reds_list.sheet_by_index(0)

pnm_primary = []
pnm_secondary = []

for i in range(red_sheet.nrows):
    check = False
    temp_arr = [red_sheet.cell_value(i, 0), red_sheet.cell_value(i, 1), red_sheet.cell_value(i, 2),
                red_sheet.cell_value(i, 3), red_sheet.cell_value(i, 4)]
    for j in range(secondary_sheet.nrows):
        if temp_arr[0] == secondary_sheet.cell_value(j, 0):
            pnm_secondary.append(temp_arr)
            check = True
            break
    if not check:
        pnm_primary.append(temp_arr)

cell_format_bold = secondary_wb.add_format({'bold': True})
cell_format_bold_p = primary_wb.add_format({'bold': True})


def write_header(worksheet):
    worksheet.write(0, 0, "PNM ID", cell_format_bold)
    worksheet.write(0, 1, "First Name", cell_format_bold)
    worksheet.write(0, 2, "Last Name", cell_format_bold)
    worksheet.write(0, 3, "Value 1", cell_format_bold)
    worksheet.write(0, 4, "Value 2", cell_format_bold)


def write_content(arr):
    row = 1
    col = 0
    for i in range(len(arr)):
        primary_ws.write(row, col, arr[0])
        primary_ws.write(row, col + 1, arr[1])
        primary_ws.write(row, col + 2, arr[2])
        primary_ws.write(row, col + 3, arr[3])
        primary_ws.write(row, col + 4, arr[4])
        row = row + 1


write_header(primary_wb)
write_header(secondary_wb)
write_content(pnm_primary)
write_content(pnm_secondary)
