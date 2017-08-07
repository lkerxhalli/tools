import sys
import os
import csv
import datetime
from re import sub
from decimal import Decimal

directory = '/Users/lkerxhalli/Documents/iris/python/aug1/'
csvinputfile = directory + 'AP Payments 1.17-7.17.csv'
csvopenbills = directory + 'Open Bills 8.1.17'
csvoutputfile = directory + 'AP Payments 1.17-7.17 out.csv'
year = 2017

outdict = {}

def parseName(str):
	if str.find(' - ') <= 0:
		return None
	arr = str.split(' - ')
	strout = arr[-1]
	if strout.find(' ') <= 0:
		print str
	index = strout.index(' ') + 1
	return strout[index:]
	
def getWeekNumber(strDate):
	dt = datetime.datetime.strptime(strDate, "%m/%d/%y").date()
	return dt.isocalendar()[1]
	
def getWeekHeader(strDate):
	dt = datetime.datetime.strptime(strDate, "%m/%d/%y").date()
	year = dt.isocalendar()[0]
	weekNo = dt.isocalendar()[1]
	strDt = "{}-W{}".format(year, weekNumber)
	monday = datetime.datetime.strptime(strDt + '-1', "%Y-W%W-%w")
	sunday = monday + datetime.timedelta(days=6)
	return monday.strftime('%m/%d/%y') + ' - ' + sunday.strftime('%m/%d/%y')

def getDateRange(weekNumber):
	strDt = "{}-W{}".format(year, weekNumber)
	monday = datetime.datetime.strptime(strDt + '-1', "%Y-W%W-%w")
	sunday = monday + datetime.timedelta(days=6)
	return monday.strftime('%m/%d/%Y') + ' - ' + sunday.strftime('%m/%d/%Y')

def getNumber(strMoney):
	value = Decimal(sub(r'[^\d.]', '', strMoney))
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
	
def removeFileHeader():
	with open(csvinputfile, 'r') as fin:
	    data = fin.read().splitlines(True)
	noLinesToSkip = 0
	for line in data:
		if isLineFull(line):
			break
		else:
			noLinesToSkip += 1
	
	with open(csvinputfile, 'w') as fout:
	    fout.writelines(data[noLinesToSkip:])

def main():
	#let's remove any extra lines at the top (Titles etc)
	removeFileHeader()
	#now process csv
	with open(csvinputfile, 'rU') as s_file:
		csv_r = csv.DictReader(s_file)
		
		tmpTransaction = ''
		
		for csv_row in csv_r:
			if csv_row['Transaction']:
				tmpTransaction = parseName(csv_row['Transaction'])
				if not tmpTransaction in outdict:
					outdict[tmpTransaction] = [0]*53
			if csv_row['Bill Type'] == 'Bill Payment':
				week = getWeekNumber(csv_row['Date'])
				amount = getNumber(csv_row['Amount'])
				outdict[tmpTransaction][week] += amount
	
	
	headers = ['Name']
	for i in range(1,52):
		headers.append(getDateRange(i))
		
	with open(csvoutputfile, 'w') as wfile:
		csvw = csv.writer(wfile, dialect='excel')
		csvw.writerow(headers)
		for key in outdict:
			outList = outdict[key]
			outList[0] = key
			csvw.writerow(outList)
			
		# csvw.writerows(lstCsvOut)
	
	
	
	
	

if __name__ == '__main__':
	main()