### Auth: Lorenc Kerxhalli
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
import re
from re import sub
from decimal import Decimal

directory = '/Users/lkerxhalli/Documents/iris/sep11/'
csvinputfile = directory + 'NPS payment history.csv'
csvopenbills = directory + 'NPS Open bills.csv'
csvoutputfile = directory + 'NPS payment history out.csv'

hName = 'Name'
hBill = 'Bills > 30' # change agingDays below if you change the number here as well
hAvg = 'Avg.'
hAvgAdjust = 'Avg. Adjusted'
extraHeaders = 4 #headers that are not weeks
agingDays = 30 #number of recent days to ignore from open bills


# TODO continue
# def isStringCompany(str):
	


def parseName(str):
	p = re.compile('V[0-9]{5}')
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
	else:
		return datetime.datetime.strptime(strDate, "%m/%d/%Y").date()
	
def getWeekHeader(dt):
	year = dt.isocalendar()[0]
	weekNo = dt.isocalendar()[1]
	strDt = "{}-W{}".format(year, weekNo)
	monday = datetime.datetime.strptime(strDt + '-1', "%Y-W%W-%w")
	sunday = monday + datetime.timedelta(days=6)
	return monday.strftime('%m/%d/%y') + ' - ' + sunday.strftime('%m/%d/%y')

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
	print ('-- Start --')
	print ('-- Clean csvs --')
	#let's remove any extra lines at the top (Titles etc)
	removeFileHeader()
	
	outdict = {}
	
	firstDate = getDate('01/01/40') #initialize them
	lastDate = getDate('01/01/10')
	
	weeks = 0 # number of weeks
	
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
	print ('Number of weeks: {}'.format(weeks))
	
	#read open bills csv
	print ('-- reading open bills --')
	openBillsDict = {}
	with open(csvopenbills, 'rU') as s_file:
		csv_r = csv.DictReader(s_file)
		for csv_row in csv_r:
			delta = datetime.date.today() - getDate(csv_row['Date Due'])
			if(delta.days > 30):
				vendor = parseName(csv_row['Vendor'])
				amount = getNumber(csv_row['Amount Due'])
				if vendor in openBillsDict:
					openBillsDict[vendor] += amount
				else:
					openBillsDict[vendor] = amount
	
	#add open bills to the main dictionary
	for vendor in openBillsDict:
		if vendor in outdict:
			outdict[vendor][hBill] = openBillsDict[vendor]
	
	#calculate and add averages
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