import xlrd
import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('recruitment2020FE6.xlsx')
worksheet = workbook.add_worksheet()


def myFunct(test):
    return test[0]


# Give the location of the file
loc = r"C:\Users\arhon\Google Drive\Recruitment 2020\raw_scores_report_e6.xlsx"

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

pnm_sort_arr = []

for i in range(sheet.nrows):
    temp_arr = [sheet.cell_value(i, 0), sheet.cell_value(i, 1), sheet.cell_value(i, 2), sheet.cell_value(i, 5)]
    pnm_sort_arr.append(temp_arr)

pnm_sort_arr.sort(key=myFunct)
pnm_final_arr = []
for i in range(0, sheet.nrows):
    temp_arr = pnm_sort_arr[i]
    for j in range(i + 1, sheet.nrows):
        if pnm_sort_arr[i][0] == pnm_sort_arr[j][0]:
            temp_arr.append(pnm_sort_arr[j][3])
        else:
            break
    if len(pnm_final_arr) > 0:
        if pnm_final_arr[len(pnm_final_arr) - 1][0] != temp_arr[0]:
            pnm_final_arr.append(temp_arr)
    else:
        pnm_final_arr.append(temp_arr)

# print("Number of reds: " + str(num_scew))
# print("Number of green: " + str(num_not))
# print("Number with 2 votes: " + str(double))

cell_format_bold = workbook.add_format({'bold': True})

worksheet.write(0, 0, "PNM ID", cell_format_bold)
worksheet.write(0, 1, "First Name", cell_format_bold)
worksheet.write(0, 2, "Last Name", cell_format_bold)
worksheet.write(0, 3, "Value 1", cell_format_bold)
worksheet.write(0, 4, "Value 2", cell_format_bold)
worksheet.write(0, 5, "Value 3", cell_format_bold)
row = 1
col = 0

# Iterate over the data and write it out row by row.
cell_format_red = workbook.add_format({'bg_color': 'red'})
cell_format_green = workbook.add_format({'bg_color': 'green'})
cell_format_yellow = workbook.add_format({'bg_color': 'yellow'})

num_skew = 0
double = 0
num_ones = 0
for i in range(len(pnm_final_arr)):
    pnm = pnm_final_arr[i][0]
    first = pnm_final_arr[i][1]
    last = pnm_final_arr[i][2]
    value1 = pnm_final_arr[i][3]
    isBad = 0
    if len(pnm_final_arr[i]) > 5:
        value2 = pnm_final_arr[i][4]
        value3 = pnm_final_arr[i][5]
        double = double + 1
        if (abs(value1 - value2 > 1)) or (abs(value1 - value3) > 1):
            isBad = 1
            num_skew = num_skew + 1
    elif len(pnm_final_arr[i]) > 4:
        value1 = pnm_final_arr[i][3]
        value2 = pnm_final_arr[i][4]
        double = double + 1
        if abs(value1 - value2) > 1:
            isBad = 1
            num_skew = num_skew + 1
    else:
        isBad = 3
        num_ones = num_ones + 1
    if isBad == 1:
        worksheet.write(row, col, pnm, cell_format_red)
        worksheet.write(row, col + 1, first, cell_format_red)
        worksheet.write(row, col + 2, last, cell_format_red)
        worksheet.write(row, col + 3, value1, cell_format_red)
        worksheet.write(row, col + 4, value2, cell_format_red)
        if len(pnm_final_arr[i]) > 5:
            worksheet.write(row, col + 5, value3, cell_format_red)
    elif isBad == 0:
        worksheet.write(row, col, pnm, cell_format_green)
        worksheet.write(row, col + 1, first, cell_format_green)
        worksheet.write(row, col + 2, last, cell_format_green)
        worksheet.write(row, col + 3, value1, cell_format_green)
        worksheet.write(row, col + 4, value2, cell_format_green)
        if len(pnm_final_arr[i]) > 5:
            worksheet.write(row, col + 5, value3, cell_format_green)
    elif isBad == 3:
        worksheet.write(row, col, pnm, cell_format_yellow)
        worksheet.write(row, col + 1, first, cell_format_yellow)
        worksheet.write(row, col + 2, last, cell_format_yellow)
        worksheet.write(row, col + 3, value1, cell_format_yellow)
    row += 1

worksheet.write(row, 0, "Number of Red: " + str(num_skew), cell_format_bold)
worksheet.write(row, 1, "Total 2s: " + str(double), cell_format_bold)
worksheet.write(row, 2, "Number of Ones: " + str(num_ones), cell_format_bold)
worksheet.write(row, 2, "Total: : " + str(double + num_ones), cell_format_bold)


# Write a total using a formula.
workbook.close()
