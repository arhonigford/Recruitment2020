import xlrd
import xlsxwriter

# Create a workbook and add a worksheet.
primary_wb = xlsxwriter.Workbook('Relist_PNMs.xlsx')
primary_ws = primary_wb.add_worksheet()

# Give the location of the file
loc_reds = r"C:\Users\arhon\Google Drive\Recruitment Data\More_Than_One_Apart.xlsx"
reds_list = xlrd.open_workbook(loc_reds)
red_sheet = reds_list.sheet_by_index(0)

pnm_list = []


for i in range(red_sheet.nrows):
    check = False
    temp_arr = [red_sheet.cell_value(i, 0), red_sheet.cell_value(i, 1), red_sheet.cell_value(i, 2),
                red_sheet.cell_value(i, 3), red_sheet.cell_value(i, 4)]
    if i == 0:
        pnm_list.append(temp_arr)
    elif abs(temp_arr[3] - temp_arr[4]) > 1:
        pnm_list.append(temp_arr)

cell_format_bold = primary_wb.add_format({'bold': True})

row = 0
col = 0

for i in range(len(pnm_list)):
    primary_ws.write(row, col, pnm_list[i][0])
    primary_ws.write(row, col + 1, pnm_list[i][1])
    primary_ws.write(row, col + 2, pnm_list[i][2])
    primary_ws.write(row, col + 3, pnm_list[i][3])
    primary_ws.write(row, col + 4, pnm_list[i][4])
    row = row + 1

primary_ws.write(row, col, "Total: " + str(len(pnm_list)))

primary_wb.close()

