# Make buttle of Moodle user from CVS file
# Author: Nguyen Thanh Tuan
# Version 1.0

import csv
#config
filename = 'onthi1'
coursename = 'cccnttcb'
header = ['username','password','firstname','lastname','email','course1']

#########################
inFile = open(filename+'.csv')
inReader = csv.reader(inFile, delimiter=',')
data = list(inReader)
outFile = open(filename + '-account.csv', 'w', newline='', encoding='utf-8')
outWriter = csv.writer(outFile)
outWriter.writerow(header)
print('Number of row: ',len(data))
for row in range(len(data)):
    username = 'ab' + data[row][0]
    firstname = data[row][0]
    lastname = data[row][1] + data[row][2]
    password = data[row][4].replace('/','')
    email = username + '@email.com'
    outWriter.writerow([username,password,firstname,lastname,email,coursename])
print(row + 1,'rows have been writen\n');