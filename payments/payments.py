### Creates a estimated payment schedule per week
### input: 
###       1. AP Payment report exported from Netsuite
###       2. Open Bills report exported from Netsuite
###
### output:
###      Estimated payment schedule per vendor per week
###


import sys
import os
import csv
import datetime
from  datetime import timedelta
from re import sub
from decimal import Decimal

directory = '/Users/lkerxhalli/Documents/iris/aug1/'
csvinputfile = directory + 'AP Payments 1.17-7.17.csv'
csvopenbills = directory + 'Open Bills 8.1.17.csv'
csvoutputfile = directory + 'AP Payments 1.17-7.17 out.csv'
year = 2017

hName = 'Name'
hBill = 'Bills > 30'
hAvg = 'Avg.'
hAvgAdjust = 'Avg. Adjusted'
extraHeaders = 4 #headers that are not weeks


def parseName(str):
	if str.find(' - ') <= 0:
		return None
	arr = str.split(' - ')
	strout = arr[-1]
	if strout.find(' ') <= 0:
		print str
	index = strout.index(' ') + 1
	return strout[index:]
	
def parseNameFromBill(str):
	index = str.index(' ') + 1
	return str[index:]
	
def getDate(strDate):
	return datetime.datetime.strptime(strDate, "%m/%d/%y").date()
	
def getWeekNumber(strDate):
	dt = getDate(strDate)
	return dt.isocalendar()[1]
	
def getWeekHeader(dt):
	year = dt.isocalendar()[0]
	weekNo = dt.isocalendar()[1]
	strDt = "{}-W{}".format(year, weekNo)
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
			break
		else:
			noLinesToSkip += 1
	
	with open(csvinputfile, 'w') as fout:
	    fout.writelines(data[noLinesToSkip:])
		
	# clean open bills
	with open(csvopenbills, 'rU') as fin:
		data = fin.read().splitlines(True)
	outData = []
	for line in data:
		if isOpenBillLine(line):
			outData.append(line)
	
	with open(csvopenbills, 'w') as fout:
	    fout.writelines(outData)	
		
def generateHeader(firstDate, lastDate):
	currentDate = firstDate
	header = [hName]
	while currentDate < lastDate:
		header.append(getWeekHeader(currentDate))
		currentDate = currentDate + timedelta(days=7)
	#take care of last date
	lastWeekHeader = getWeekHeader(lastDate)
	if header[-1] != lastWeekHeader:
		header.append(lastWeekHeader)
	
	header.append(hBill)
	header.append(hAvg)
	header.append(hAvgAdjust)
	return header


def main():
	print '-- Start --'
	print '-- Clean csvs --'
	#let's remove any extra lines at the top (Titles etc)
	removeFileHeader()
	
	outdict = {}
	
	firstDate = getDate('01/01/40') #initialize them
	lastDate = getDate('01/01/10')
	
	weeks = 0 # number of weeks
	
	print '-- reading AP --'
	#now read payment csv
	with open(csvinputfile, 'rU') as s_file:
		csv_r = csv.DictReader(s_file)
		
		tmpTransaction = '' # transaction string includes the vendor name
		
		for csv_row in csv_r:
			if csv_row['Transaction']:
				tmpTransaction = parseName(csv_row['Transaction'])
				if not tmpTransaction in outdict:
					outdict[tmpTransaction] = {}
			if csv_row['Bill Type'] == 'Bill Payment':
				dt = getDate(csv_row['Date'])
				week = getWeekHeader(dt)
				if dt > lastDate:
					lastDate = dt
				if dt < firstDate:
					firstDate = dt
				amount = getNumber(csv_row['Amount'])
				if not week in outdict[tmpTransaction]:
					outdict[tmpTransaction][week] = amount
				else:
					outdict[tmpTransaction][week] += amount
	
	header = generateHeader(firstDate, lastDate)
	weeks = len(header) - extraHeaders
	print 'Number of weeks: {}'.format(weeks)
	
	#read open bills csv
	print '-- reading open bills --'
	openBillsDict = {}
	with open(csvopenbills, 'rU') as s_file:
		csv_r = csv.DictReader(s_file)
		for csv_row in csv_r:
			delta = datetime.date.today() - getDate(csv_row['Date Due'])
			if(delta.days > 30):
				vendor = parseNameFromBill(csv_row['Vendor'])
				amount = getNumber(csv_row['Amount Due'])
				if vendor in openBillsDict:
					openBillsDict[vendor] += amount
				else:
					openBillsDict[vendor] = amount
	
	for vendor in openBillsDict:
		if vendor in outdict:
			outdict[vendor][hBill] = openBillsDict[vendor]
	
	for vendor in outdict:
		avg = 0
		for key in outdict[vendor]:
			if key != hName and key != hBill:
				avg += outdict[vendor][key]
		if weeks > 0:
			outdict[vendor][hAvg] = round(avg/weeks, 2)
			if hBill in outdict[vendor]:
				outdict[vendor][hAvgAdjust] = round((avg + outdict[vendor][hBill])/weeks, 2)
			else:
				outdict[vendor][hAvgAdjust] = outdict[vendor][hAvg]
			
			
	print '-- Completing Calculations --'
		
	#TNE for the names

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
	
	print '-- finito --'
	
	
	
	
	

if __name__ == '__main__':
	main()