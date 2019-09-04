import openpyxl
import csv
#config
filename = 'nmcntt19'
coursename = 'nmcntt19'
header = ['username','password','firstname','lastname','email','course1']
firstrow = 2

#########################
wb = openpyxl.load_workbook(filename + '.xlsx')
print(wb.get_sheet_names)
outFile = open(filename + '-accounts.csv', 'w', newline='', encoding='utf-8')
outWriter = csv.writer(outFile)

sheet = wb.active
print(sheet)
outWriter.writerow(header)
for row in range(firstrow, sheet.max_row + 1):
    username = str(sheet['B' + str(row)].value)
    firstname = sheet['C' + str(row)].value
    lastname = sheet['D' + str(row)].value
    # print(sheet['E' + str(row)].value)
    password = str(sheet['E' + str(row)].value).replace('-','')
    email = username + '@email.com'
    print([username,password,firstname,lastname,email,coursename])
    outWriter.writerow([username,password,firstname,lastname,email,coursename])
