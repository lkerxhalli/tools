### Auth: Lorenc Kerxhalli
### Creates a estimated payment schedule per month
### input:
###       1. Invoices Payment report exported from Netsuite
###
### output:
###      Estimated payment schedule per vendor per month
###


import sys
import os
import csv
import datetime
from  datetime import timedelta
import calendar
from dateutil.relativedelta import *
import re
from re import sub
from decimal import Decimal

directory = '/Users/lkerxhalli/Documents/iris/jun14/'
name = 'A_PPaymentHistorybyBill968'
csvinputfile = directory + name + '.csv'
csvoutputfile = directory + name +  ' out.csv'

hName = 'Name'
hBill = 'Bills > 30' # change agingDays below if you change the number here as well
hAvg = 'Avg.'
hAvgAdjust = 'Avg. Adjusted'
extraHeaders = 4 #headers that are not months
agingDays = 30 #number of recent days to ignore from open bills


# TODO continue
# def isStringCompany(str):



def parseName(str):
    p = re.compile(' V[0-9]{5}')
    m = p.search(str)
    if(m):
        index = m.end() + 1
        return str[index:]
    else:
        return None




def getDate(strDate):
    # check if strdate year part is 2 digit or 4
    lIndex = strDate.rfind('/') + 1 # index of first digit
    strYear = strDate[lIndex:]
    if(len(strYear) == 2):
        return datetime.datetime.strptime(strDate, "%m/%d/%y").date()
    elif (len(strYear) == 4):
        return datetime.datetime.strptime(strDate, "%m/%d/%Y").date()
    else:
        return None

def getMonthHeader(dt):
    year = dt.isocalendar()[0]
    month = dt.month
    strDt = "{} {}".format(calendar.month_abbr[month], year)
    return strDt

def getNumber(strMoney):
    value = Decimal(sub(r'[^\d.]', '', strMoney))
    if strMoney.find('(') > -1:
        value = value * -1
    return value

def isLineFull(strLine):
    arrLine = strLine.split(',')
    countFields = 0
    for field in arrLine:
        if field:
            countFields += 1
    if countFields > 4:
        return True
    else:
        return False

def isOpenBillLine(strLine):
    arrLine = strLine.split(',')
    if len(arrLine) > 5 and arrLine[5]:
        return True
    else:
        return False

def removeFileHeader():
    with open(csvinputfile, 'r') as fin:
        data = fin.read().splitlines(True)
    noLinesToSkip = 0
    for line in data:
        if isLineFull(line):
            # hdr = line.split(',')
            # strLine = ''
            # for item in hdr:
            #     strLine += item.strip() + ','
            # strLine = strLine[:-1]
            # data[noLinesToSkip] = strLine
            break
        else:
            noLinesToSkip += 1

    with open(csvinputfile, 'w') as fout:
        fout.writelines(data[noLinesToSkip:])


def generateHeader(firstDate, lastDate):
    currentDate = firstDate
    header = [hName]
    while currentDate < lastDate:
        header.append(getMonthHeader(currentDate))
        currentDate = currentDate + relativedelta(months=+1)
    #take care of last date
    lastMonthHeader = getMonthHeader(lastDate)
    if header[-1] != lastMonthHeader:
        header.append(lastMonthHeader)

    header.append(hBill)
    header.append(hAvg)
    header.append(hAvgAdjust)
    return header


def main():
    print ('-- Start --')
    print ('-- Clean csvs --')
    #let's remove any extra lines at the top (Titles etc)
    removeFileHeader()

    outdict = {}

    firstDate = getDate('01/01/40') #initialize them
    lastDate = getDate('01/01/10')

    months = 0 # number of months

    print ('-- reading AP --')
    #now read payment csv
    with open(csvinputfile, 'rU') as s_file:
        csv_r = csv.DictReader(s_file)

        tmpTransaction = '' # transaction string includes the vendor name

        for csv_row in csv_r:
            if csv_row['Transaction']:
                tmpTransaction = parseName(csv_row['Transaction'])
                if not tmpTransaction in outdict:
                    outdict[tmpTransaction] = {}
            if csv_row['Payment Type'] == 'Bill':
                dt = getDate(csv_row['Date'])
                month = getMonthHeader(dt)
                if dt > lastDate:
                    lastDate = dt
                if dt < firstDate:
                    firstDate = dt
                amount = getNumber(csv_row['Amount'])
                if not month in outdict[tmpTransaction]:
                    outdict[tmpTransaction][month] = amount
                else:
                    outdict[tmpTransaction][month] += amount

    header = generateHeader(firstDate, lastDate)
    months = len(header) - extraHeaders
    print ('Number of months: {}'.format(months))


    #calculate and add averages
    for vendor in outdict:
        avg = 0
        for key in outdict[vendor]:
            if key != hName and key != hBill:
                avg += outdict[vendor][key]
        if months > 0:
            outdict[vendor][hAvg] = round(avg/months, 2)
            if hBill in outdict[vendor]:
                outdict[vendor][hAvgAdjust] = round((avg + outdict[vendor][hBill])/months, 2)
            else:
                outdict[vendor][hAvgAdjust] = outdict[vendor][hAvg]


    print ('-- Completing Calculations --')

    #TODO: find and Sort for TNE for the names

    # for python 3 use below to prevent extra blank lines
    # with open(csvoutputfile, 'w', newline='') as wfile:
    with open(csvoutputfile, 'w') as wfile:
        csvw = csv.writer(wfile, dialect='excel')
        csvw.writerow(header)
        for key in outdict:
            if key:
                row = [key]
                for col in header[1:]:
                    if col in outdict[key]:
                        row.append(outdict[key][col])
                    else:
                        row.append(0)
                csvw.writerow(row)

    print ('-- finito --')






if __name__ == '__main__':
    main()
