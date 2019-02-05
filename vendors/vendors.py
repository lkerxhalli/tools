### Auth: Lorenc Kerxhalli
### Aggregates info by vendors
### input:
###       TransactionSearchResults360.csv
###
### output:
###      VendorSummary.csv
### ignore 21000 Accounts Payable
###


import sys
import os
import csv
import datetime
from  datetime import timedelta
import re
from re import sub
from decimal import Decimal

directory = '/Users/lkerxhalli/Documents/iris/feb4/'
csvinputfile = directory + 'TransactionSearchResults253.csv'
csvoutputfile = directory + 'VendorSummary384.csv'


def main():
    print ('--- Start ---')
    outdict = {} # key is composite of Vendor and Subsidiary
    header = ['Subsidiary', 'Vendor', 'Amount']

    with open(csvinputfile, 'rU') as s_file:
        csv_r = csv.DictReader(s_file)

        index = 0
        vendorName = ''
        maxAccount = 1
        maxTempAccount = 0
        maxServiceLine = 1
        maxTempServiceLine = 0
        maxLocation = 1
        maxTempLocation = 0
        maxDepartment = 1
        maxTempDepartment = 0
        for csv_row in csv_r:
            index += 1
            # if index > 100:
            #      break
            if csv_row['Name'] != '':
                vendorName = csv_row['Name']
            if True: #vendorName == 'V19464 4900 Perry Highway Management LLC':
                subsidiary = csv_row['Subsidiary']
                account = str(csv_row['Account'].split()[0])
                amount = abs(float(csv_row['Amount']))
                type = csv_row['Type']
                serviceLine = csv_row['Service Line']
                location = csv_row['Location']
                department = csv_row['Department']

                # print "%s| %s| %s| %s| %s" % (vendorName, csv_row['Account'], serviceLine, location, department)

                key = vendorName + subsidiary
                if not key in outdict:
                    outdict[key] = {}
                    if maxTempAccount > maxAccount:
                        maxAccount = maxTempAccount
                    maxTempAccount = 0
                    if maxTempServiceLine > maxServiceLine:
                        maxServiceLine = maxTempServiceLine
                    maxTempServiceLine = 0
                    if maxTempDepartment > maxDepartment:
                        maxDepartment = maxTempDepartment
                    maxTempDepartment = 0
                    if maxTempLocation > maxLocation:
                        maxLocation = maxTempLocation
                    maxTempLocation = 0
                if account != '21000' and type == 'Bill':
                    outdict[key]['Subsidiary'] = subsidiary
                    outdict[key]['Vendor'] = vendorName
                    if 'Amount' in outdict[key]:
                        outdict[key]['Amount'] += amount
                    else:
                        outdict[key]['Amount'] = amount
                    if not 'Accounts' in outdict[key]:
                        outdict[key]['Accounts'] = set()
                    outdict[key]['Accounts'].add(account)
                    maxTempAccount = len(outdict[key]['Accounts'])
                    if not 'ServiceLines' in outdict[key]:
                        outdict[key]['ServiceLines'] = set()
                    outdict[key]['ServiceLines'].add(serviceLine)
                    maxTempServiceLine = len(outdict[key]['ServiceLines'])

                    if not 'Departments' in outdict[key]:
                        outdict[key]['Departments'] = set()
                    outdict[key]['Departments'].add(department)
                    maxTempDepartment = len(outdict[key]['Departments'])

                    if not 'Locations' in outdict[key]:
                        outdict[key]['Locations'] = set()
                    outdict[key]['Locations'].add(location)
                    maxTempLocation = len(outdict[key]['Locations'])


    print maxAccount
    for i in range(1, maxAccount + 1):
        header.append('Account {}'.format(i))
    for i in range(1, maxServiceLine + 1):
        header.append('Service Line {}'.format(i))
    for i in range(1, maxDepartment + 1):
        header.append('Department {}'.format(i))
    for i in range(1, maxLocation + 1):
        header.append('Location {}'.format(i))

    with open(csvoutputfile, 'w') as wfile:
        csvw = csv.writer(wfile, dialect='excel')
        csvw.writerow(header)
        for key in outdict:
            if 'Subsidiary' in outdict[key]:
                lastIndex = 0
                row = [outdict[key]['Subsidiary'], outdict[key]['Vendor'], outdict[key]['Amount']]
                for acc in outdict[key]['Accounts']:
                    row.append(acc)
                    lastIndex += 1
                tempIndex = lastIndex
                for i in range(tempIndex, maxAccount):
                    row.append('')
                    lastIndex += 1
                for sline in outdict[key]['ServiceLines']:
                    row.append(sline)
                    lastIndex += 1
                tempIndex = lastIndex
                for i in range(tempIndex, maxAccount+maxServiceLine):
                    row.append('')
                    lastIndex += 1
                for dep in outdict[key]['Departments']:
                    row.append(dep)
                    lastIndex += 1
                tempIndex = lastIndex
                for i in range(tempIndex, maxAccount+maxServiceLine+maxDepartment):
                    row.append('')
                    lastIndex += 1
                for loc in outdict[key]['Locations']:
                    row.append(loc)
                csvw.writerow(row)



if __name__ == '__main__':
    main()
