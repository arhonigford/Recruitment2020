import xlrd
import xlsxwriter

# Create a workbook and add a worksheet.
primary_wb = xlsxwriter.Workbook('scewedPNMs.xlsx')
primary_ws = primary_wb.add_worksheet('Primary')
secondary_ws = primary_wb.add_worksheet('Secondary')

# Give the location of the file
loc_reds = r"C:\Users\arhon\Google Drive\Recruitment Data\More_Than_One_Apart.xlsx"
reds_list = xlrd.open_workbook(loc_reds)
red_sheet = reds_list.sheet_by_index(0)

loc_secondary = r"C:\Users\arhon\Google Drive\Recruitment Data\secondary_list.xlsx"
secondary_list = xlrd.open_workbook(loc_secondary)
secondary_sheet = secondary_list.sheet_by_index(0)

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


cell_format_bold = primary_wb.add_format({'bold': True})

row = 0
col = 0

for i in range(len(pnm_primary)):
    primary_ws.write(row, col, pnm_primary[i][0])
    primary_ws.write(row, col + 1, pnm_primary[i][1])
    primary_ws.write(row, col + 2, pnm_primary[i][2])
    primary_ws.write(row, col + 3, pnm_primary[i][3])
    primary_ws.write(row, col + 4, pnm_primary[i][4])
    row = row + 1
primary_ws.write(row, col, "Total: " + str(len(pnm_primary) - 1))

row = 0
secondary_ws.write(row, col, pnm_primary[0][0])
secondary_ws.write(row, col + 1, pnm_primary[0][1])
secondary_ws.write(row, col + 2, pnm_primary[0][2])
secondary_ws.write(row, col + 3, pnm_primary[0][3])
secondary_ws.write(row, col + 4, pnm_primary[0][4])
row = 1
col = 0
for i in range(len(pnm_secondary)):
    secondary_ws.write(row, col, pnm_secondary[i][0])
    secondary_ws.write(row, col + 1, pnm_secondary[i][1])
    secondary_ws.write(row, col + 2, pnm_secondary[i][2])
    secondary_ws.write(row, col + 3, pnm_secondary[i][3])
    secondary_ws.write(row, col + 4, pnm_secondary[i][4])
    row = row + 1

secondary_ws.write(row, col, "Total: " + str(len(pnm_secondary)))


primary_wb.close()

