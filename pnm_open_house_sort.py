import xlrd
import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('recruitment2020test6.xlsx')
worksheet = workbook.add_worksheet()

def myFunct(test):
    return test[0]
# Give the location of the file
loc = (r"C:\Users\arhon\Google Drive\Recruitment 2020\raw_scores_report_test.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

pnm_sort_arr = []

for i in range(sheet.nrows):
    temp_arr = [sheet.cell_value(i, 0), sheet.cell_value(i, 1), sheet.cell_value(i, 2), sheet.cell_value(i, 5)]
    pnm_sort_arr.append(temp_arr)

pnm_sort_arr.sort(key=myFunct)
num_scew = 0
num_not = 0
final_pnm_arr = []
print(sheet.nrows)
double = 0
# temp_arr = [pnm id, first, last, value1]
for i in range(sheet.nrows - 1):
    if pnm_sort_arr[i][0] == pnm_sort_arr[i + 1][0]:
        double += 1
        temp_arr = pnm_sort_arr[i]
        temp_arr.append(pnm_sort_arr[i + 1][3])
        if (pnm_sort_arr[i][3] - pnm_sort_arr[i+1][3]) > 1:
            temp_arr.append('red')
            final_pnm_arr.append(temp_arr)
            num_scew += 1
        else:
            temp_arr.append('green')
            num_not += 1
            final_pnm_arr.append(temp_arr)
        #print(pnm_sort_arr[i])
print("Number of reds: " + str(num_scew))
print("Number of green: " + str(num_not))
print("Number with 2 votes: " + str(double))

cell_format_bold= workbook.add_format({'bold': True})

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


for pnm, first, last, value1, value2, color in (final_pnm_arr):
    if color.find('red'):
        worksheet.write(row, col, pnm, cell_format_green)
        worksheet.write(row, col + 1, first, cell_format_green)
        worksheet.write(row, col + 2, last, cell_format_green)
        worksheet.write(row, col + 3, value1, cell_format_green)
        worksheet.write(row, col + 4, value2, cell_format_green)
        worksheet.write(row, col + 5, color, cell_format_green)
    else:
        worksheet.write(row, col, pnm, cell_format_red)
        worksheet.write(row, col + 1, first, cell_format_red)
        worksheet.write(row, col + 2, last, cell_format_red)
        worksheet.write(row, col + 3, value1, cell_format_red)
        worksheet.write(row, col + 4, value2, cell_format_red)
        worksheet.write(row, col + 5, color, cell_format_red)
    row += 1

worksheet.write(row, 0, "Number of Red: " + str(num_scew), cell_format_bold)
worksheet.write(row, 1, "Number of Green: " + str(num_not), cell_format_bold)
worksheet.write(row, 2, "Total 2s: " + str(double), cell_format_bold)

# Write a total using a formula.
workbook.close()


