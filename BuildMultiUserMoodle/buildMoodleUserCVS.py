# Make buttle of Moodle user from CVS file
# Author: Nguyen Thanh Tuan
# Version 1.0

import csv
#config
filename = 'test'
coursename = 'testcourse'
header = ['username','password','firstname','lastname','email','course1']

#########################
inFile = open(filename+'.cvs')
inReader = csv.reader(inFile, delimiter='\t')
data = list(inReader)
outFile = open(filename + '-account.csv', 'w', newline='', encoding='utf-8')
outWriter = csv.writer(outFile)
outWriter.writerow(header)
print('Number of row: ',len(data))
for row in range(len(data)):
    username = data[row][1]
    firstname = data[row][2]
    lastname = data[row][3] + ' ' + data[row][5]
    password = data[row][4].replace('-','')
    email = username + '@email.com'
    outWriter.writerow([username,password,firstname,lastname,email,coursename])
print(row + 1,'rows have been writen\n');