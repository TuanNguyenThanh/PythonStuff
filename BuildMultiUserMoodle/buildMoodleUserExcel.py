import openpyxl
import csv
#config
filename = 'example'
coursename = 'course'
header = ['username','password','firstname','lastname','email','course1']

#########################
wb = openpyxl.load_workbook(filename + '.xlsx')
outFile = open(filename + '.csv', 'w', newline='', encoding='utf-8')
outWriter = csv.writer(outFile)

sheet = wb.active
outWriter.writerow(header)
for row in range(2, sheet.max_row + 1):
    username = sheet['B' + str(row)].value
    firstname = sheet['C' + str(row)].value
    lastname = sheet['D' + str(row)].value + ' ' + sheet['F' + str(row)].value
    password = sheet['E' + str(row)].value.replace('-','')
    email = username + '@email.com'
    outWriter.writerow([username,password,firstname,lastname,email,coursename])
