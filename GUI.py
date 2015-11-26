## 2015/5/18 @ 1500

import os
import sys
import csv
import subprocess
import time
from operator import itemgetter
import time
import Tkinter
from Tkinter import *
import tkFileDialog
from tkFileDialog import *
import tkMessageBox
from tkMessageBox import *

#### Part Validation Worker ####

class PartValidation:
	def addToBox(self, boxNum, entry):
		self.box[boxNum].append(entry)
		self.box[boxNum].sort()
		print "Added " + entry[0] + " - " + entry[1] + " to box " + str(boxNum) + '.'

	def addBox(self):
		self.box.append([])

	def scanIn(self, data, partNum, repairNum, boxNum):
		if len(data) > 0:
			foundEntries = []
			if partNum.startswith('76-'):
				partNum = '0' + partNum
			checked = False
			found = False
			entry = 0
			foundEntries = []
			while (checked == False):
				reason = "No Match\nCheck the \"Moved today\"\nIf not there, manually run transaction history."
				if (repairNum == data[entry][1]):
					print "Found repair number..."
					if (partNum == data[entry][0]):
						print "Found part number under repair number..."
						gotAll = True
						if ((data[entry][3] - data[entry][4]) >= 2):
							gotAll = tkMessageBox.askokcancel("Multi POS'd", str(data[entry][5]) + " " + data[entry][0] + "\'s were POS'd, do you have all of them?")
						if (partNum.startswith('661') and gotAll == True):
							print "checking 661 part..."
							if ( (data[entry][3] - data[entry][4]) == data[entry][5] and data[entry][5] == data[entry][6] and data[entry][5] > 0):
								#if self.printOut == True:
									#print "\t#### Valid repair as " + data[entry][2] + " ####"
								##self.validParts.append(data[entry])
								self.addToBox(boxNum, data[entry])
								data.pop(entry)
								foundEntries.append(entry)
								entry -= 1
								found = True
							else :
								reason = "Found " + data[entry][0] + " - " + data[entry][1] + ", but it is unbalanced in movements.\n"
								reason += "Allocation(s): " + str(data[entry][3]) + "\n"
								reason += "Deallocations(s): " + str(data[entry][4]) + "\n"
								reason += "Sale(s): " + str(data[entry][5]) + "\n"
								reason += "RTW(s): " + str(data[entry][6]) + "\n"
								reason += "Manually run transaction history for this repair and investigate.**"
								found = True
						elif (gotAll == True):
							print "Checking consumable part..."
							if ( (data[entry][3] - data[entry][4]) == data[entry][5] and data[entry][5] > 0):
								##self.addToBox(boxNum, data[entry])
								data.pop(entry)
								foundEntries.append(-2)
								entry -= 1
								found = True
							else :
								reason = "Found " + data[entry][0] + " - " + data[entry][1] + ", but it is unbalanced in movements.\n"
								reason += "Allocation(s): " + str(data[entry][3]) + "\n"
								reason += "Deallocations(s): " + str(data[entry][4]) + "\n"
								reason += "Sale(s): " + str(data[entry][5]) + "\n"
								reason += "RTW(s): " + str(data[entry][6]) + "\n"
								reason += "Manually run transaction history for this repair and investigate.**"
								found = True
				entry += 1
				if (entry == len(data)):
					checked = True
			if found == False:
				return [ -1, reason ]
			else:
				print foundEntries
				return foundEntries
		else:
			return [ -1, "No Data"]

	def buildMovementTable(self):
		## data = [ repairNum, partNum, partDescription, movementCode, stockType, repairNum, timeStamp ]
		## tempData = [ repairNum, partNum, partDescription, # of 344, # of 343, # of sold (range(947,961) & 251), # of returns 252, timeStamp ]

		## look for repairs returned at POS
		self.returnedAtPOS = []
		for index in range(0, len(self.data)):
			if (self.data[index][3] == "252" and self.data[index][4] == "Block"):
					self.returnedAtPOS.append(self.data[index])
					if self.printOut == True:
						print "Found repair returned by manager."
						print self.data[index]

		## Build array of movements for repair numbers and part numbers
		# [ Repair Number, Part Number, Part Description, Movement, Stock Type, TimeStamp, Location Code ]
		tempData = []
		tempData.append([ "", "", "", 0, 0, 0, 0, ""])
		for movement in self.data:
			found = False
			for index in range(0, len(tempData)):
				if ((movement[0] == tempData[index][1]) and (movement[1] == tempData[index][0])):
					if (movement[3] == "344"):
						tempData[index][3] += 1
						tempData[index][7] = movement[5]
					elif (movement[3] == "343"):
						tempData[index][4] += 1
						tempData[index][7] = movement[5]
					elif (int(movement[3]) in self.reversalMovements):
						tempData[index][5] -= 1
						tempData[index][7] = movement[5]
					elif (int(movement[3]) in self.saleMovements):
						tempData[index][5] += 1
						tempData[index][7] = movement[5]
					elif (movement[3] == "251" and movement[6] == "001"):
						tempData[index][5] += 1
						tempData[index][7] = movement[5]
					elif (movement[3] == "251" and movement[6] == "002"):
						tempData[index][6] -= 1
						tempData[index][7] = movement[5]
					elif (movement[3] == "252" and movement[6] == "001"):
						tempData[index][5] -= 1
						tempData[index][7] = movement[5]
					elif (movement[3] == "252" and movement[6] == "002"):
						tempData[index][6] += 1
						tempData[index][7] = movement[5]
					found = True
			if found == False:
				tempData.append([ movement[1], movement[0], movement[2], 0, 0, 0, 0, ""])
				index = len(tempData)-1
				if (movement[3] == "344"):
					tempData[index][3] += 1
					tempData[index][7] = movement[5]
				elif (movement[3] == "343"):
					tempData[index][4] += 1
					tempData[index][7] = movement[5]
				elif (int(movement[3]) in self.reversalMovements):
					tempData[index][5] -= 1
					tempData[index][7] = movement[5]
				elif (int(movement[3]) in self.saleMovements):
					tempData[index][5] += 1
					tempData[index][7] = movement[5]
				elif (movement[3] == "251" and movement[6] == "001"):
					tempData[index][5] += 1
					tempData[index][7] = movement[5]
				elif (movement[3] == "251" and movement[6] == "002"):
					tempData[index][6] -= 1
					tempData[index][7] = movement[5]
				elif (movement[3] == "252" and movement[6] == "001"):
					tempData[index][5] -= 1
					tempData[index][7] = movement[5]
				elif (movement[3] == "252" and movement[6] == "002"):
					tempData[index][6] += 1
					tempData[index][7] = movement[5]
		tempData.pop(0)
		tempData.sort(key=itemgetter(7,0))
		self.data = tempData

	def readFile(self):
		if self.printOut == True:
			print "Opening Spreadsheet... (may take a while)"
		error = False
		try:
			workbook = xlrd.open_workbook(self.dataFile)
		except IOError:
			if self.printOut == True:
				print "Could not open file! Please close Excel!"
			error = True
		if (error == False):
			worksheet = workbook.sheet_by_name('Sheet1')
			num_rows = worksheet.nrows - 1
			num_cells = worksheet.ncols - 1
			curr_row = -1
			self.data = []
			gotTitle = False
			while curr_row < num_rows:
				curr_row += 1
				if self.printOut == True:
					sys.stdout.write("Reading rows: %d%%   \r" % ((curr_row/num_rows)*100) )
					sys.stdout.flush()
				curr_cell = -1
				exclude = [0, 4, 7, 11, 13, 14, 15]
				row_values = []
				while curr_cell < num_cells:
					curr_cell += 1
					# Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
					if (curr_cell not in exclude):
						if worksheet.cell_type(curr_row, curr_cell) == 3:
							ms_date_number = worksheet.cell_value(curr_row, curr_cell)
							curr_cell += 1
							ms_date_number2 = worksheet.cell_value(curr_row, curr_cell)
							year, month, day, hour, minute, second = xlrd.xldate_as_tuple(ms_date_number, workbook.datemode)
							year2, month2, day2, hour2, minute2, second2 = xlrd.xldate_as_tuple(ms_date_number2, workbook.datemode)
							tempStr = str(month) + "/" + str(day) + "/" + str(year) + " " + str(hour2) + ":" + str(minute2) + ":" + str(second2)
							cell_value = time.strptime(tempStr, "%m/%d/%Y %H:%M:%S")
						else :
							cell_value = worksheet.cell_value(curr_row, curr_cell).encode('ascii','ignore')
						row_values.append(cell_value)
				if (gotTitle == False):
					gotTitle = True
				else:
					## [TimeStamp, Part Number, Part Description, Movement, Stock Type, Issueing Location, Recieving Location, Repair Number]
					row_values[2] = row_values[2].replace('"', '')
					if ((int(row_values[3]) in [ 251, 252, 344, 343 ]) or (int(row_values[3]) in self.saleMovements) or (int(row_values[3]) in self.reversalMovements)):
						if (("922-" in row_values[1]) or ("923-" in row_values[1]) or ("076-" in row_values[1]) or ("661-" in row_values[1])):
							## Add the info to the data array
							if row_values[5] != '':
								location = row_values[5]
							else:
								location = row_values[6]
							# [ Repair Number, Part Number, Part Description, Movement, Stock Type, TimeStamp, Location Code ]
							self.data.append([ row_values[7], row_values[1], row_values[2], row_values[3], row_values[4], row_values[0], location ])
		if self.printOut == True:
			print "\n"
		self.data.sort(key=itemgetter(5))
		newName = "/Users/apple/Desktop/" + time.strftime("%m-%d-%y_@%I-%M-%S-%p", self.data[len(self.data)-1][5]) + self.dataFile[self.dataFile.rfind('/'):self.dataFile.rfind('.')] +"_"+ time.strftime("%m-%d-%y_@%I-%M-%S-%p", self.data[len(self.data)-1][5]) + self.dataFile[self.dataFile.rfind('.'):]
		#print newName
		os.system("mkdir /Users/apple/Desktop/" + time.strftime("%m-%d-%y_@%I-%M-%S-%p", self.data[len(self.data)-1][5]) + '/')
		os.system("mv " + self.dataFile + " " + newName)
		self.dataFile = newName

		## Table the raw data
		self.buildMovementTable()

		print "Filterings out items from other days..."
		## Filter items from previous then selected day(s)
		tempFilter = self.dateFilter
		top = len(self.data)-1
		pos = 0
		while pos <= top:
			if self.data[pos][7] < tempFilter:
				self.otherDays.append(self.data[pos])
				self.data.pop(pos)
				pos -= 1
				top -= 1
			pos += 1

		print "Filtering today's unbalanced..."
		## Filter today's unbalanced items
		top = len(self.data)-1
		pos = 0
		while pos <= top:
			if self.data[pos][5] > 0:
				if ( self.data[pos][0].startswith('661') and ((self.data[pos][3] - self.data[pos][4]) == self.data[pos][5]) and (self.data[pos][5] == self.data[pos][6]) and (self.data[pos][5] > 0) ):
					pass
				elif ( (not self.data[pos][0].startswith('661')) and (self.data[pos][3] - self.data[pos][4]) == self.data[pos][5] and self.data[pos][5] > 0):
					pass
				else:
					self.curDayPOSdUnbalanced.append(self.data[pos])
					self.data.pop(pos)
					pos -= 1
					top -= 1
			else:
				self.curDayUnbalanced.append(self.data[pos])
				self.data.pop(pos)
				pos -= 1
				top -= 1
			pos += 1

		print "Filterings other days unbalanced."
		## Filter other day's unbalanced items
		top = len(self.otherDays)-1
		print "Other days = " + str(len(self.otherDays))
		pos = 0
		while pos <= top:
			if ( ("661" in self.otherDays[pos][0]) and ((self.otherDays[pos][3] - self.otherDays[pos][4]) == self.otherDays[pos][5]) and
				(self.otherDays[pos][5] == self.otherDays[pos][6]) and (self.otherDays[pos][5] > 0) ):
				True
			elif ( (self.otherDays[pos][3] - self.otherDays[pos][4]) == self.otherDays[pos][5] and self.otherDays[pos][5] > 0):
				True
			else:
				self.otherDaysUnbalanced.append(self.otherDays[pos])
				self.otherDays.pop(pos)
				pos -= 1
				top -= 1
			pos += 1

		self.otherDays.sort()
		self.ready = True
		return error

	def getFile_console(self):
		self.dataFile = raw_input("Drag the XSLX file here and press enter: ")
		if (self.dataFile[len(self.dataFile)-1] == ' '):
			self.dataFile = self.dataFile[0:len(dataFile)-1]
		self.dataFile = dataFile.replace('\\', '')

		ext = self.dataFile[self.dataFile.rfind('.')+1:]
		validExt = False
		if (ext == "xlsx" or ext == "XLSX"):
			print "File: " + dataFile
			error = self.readInXLSX_Allocate()
			validExt = True

	def writeCount(self):
		slotNum = 0
		outCount = self.dataFile[:self.dataFile.rfind('.')] + "_SRTW_Move.jay.csv"
		opened = True
		try:
			outFile = open(outCount, "w")
		except IOError:
			if self.printOut == True:
				print "Could not open count output file!"
			opened = False
		if (opened == True):
			for slot in self.box:
				count = []
				for entry in slot:
					exists = False
					for index in count:
						if index[0] == entry[0]:
							index[1] += int(entry[5])
							exists = True
							break
					if (exists == False):
						count.append([ entry[0] , int(entry[5]), slotNum+1 ])
				count.sort(key=itemgetter(2))
				slotNum += 1
				if len(count) > 0:
					if self.printOut == True:
						print "Writing out: " + outCount
					outFile.write("\"Part Number\",Quantity,Box\r")
					for entry in count:
						outFile.write(entry[0] + "," + str(entry[1]) + "," + str(entry[2]) + "\r")
					outFile.write('\r')
			outFile.close()

	def writeValid(self):
		slotNum = 0
		outValid = self.dataFile[:self.dataFile.rfind('.')] + "_valid.jay.csv"
		opened = True
		try:
			outFile = open(outValid, "w")
		except IOError:
			if self.printOut == True:
				print "Could not open valid output file!"
			opened = False
		if (opened == True):
			for slot in self.box:
				if self.printOut == True:
					print "Writing out: " + outValid
				outFile.write("Part Number, Repair Number, Part Description, Alloc, De-Alloc, POS, RTW, Time Stamp, Box\r")
				for entry in slot:
					timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[7])
					outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + str(entry[3]) + "," + str(entry[4]) + "," + str(entry[5]) + "," + str(entry[6]) + "," + timeStampStr + "," + str(slotNum+1) + "\r")
				outFile.write('\r')
				slotNum += 1
			outFile.close()

	def writeMissing_Balanced(self):
		self.data.sort()
		outMissing = self.dataFile[:self.dataFile.rfind('.')] + "_missing.jay.csv"
		opened = True
		if (len(self.data) > 0):
			try:
				outFile = open(outMissing, "w")
			except IOError:
				print "Could not open invalid output file!"
				opened = False
			if (opened == True):
				if self.printOut == True:
					print "Writing out: " + outMissing
				outFile.write("Part Number, Repair Number, Part Description, Alloc, De-Alloc, POS, RTW, Time Stamp\r")
				for entry in self.data:
					timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[7])
					outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + str(entry[3]) + "," + str(entry[4]) + "," + str(entry[5]) + "," + str(entry[6]) + "," + timeStampStr + "\r")
				outFile.close()

	def writeMissing_POSdUnbalanced(self):
		outMissing = self.dataFile[:self.dataFile.rfind('.')] + "_POSd_but_unbalanced.jay.csv"
		opened = True
		if (len(self.curDayPOSdUnbalanced) > 0):
			try:
				outFile = open(outMissing, "w")
			except IOError:
				print "Could not open invalid output file!"
				opened = False
			if (opened == True):
				if self.printOut == True:
					print "Writing out: " + outMissing
				outFile.write("Part Number, Repair Number, Part Description, Alloc, De-Alloc, POS, RTW, Time Stamp\r")
				for entry in self.curDayPOSdUnbalanced:
					timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[7])
					outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + str(entry[3]) + "," + str(entry[4]) + "," + str(entry[5]) + "," + str(entry[6]) + "," + timeStampStr + "\r")
				outFile.close()

	def writeMissing_Unbalanced(self):
		outMissing = self.dataFile[:self.dataFile.rfind('.')] + "_missing_but_unbalanced.jay.csv"
		opened = True
		if (len(self.curDayUnbalanced) > 0):
			try:
				outFile = open(outMissing, "w")
			except IOError:
				print "Could not open invalid output file!"
				opened = False
			if (opened == True):
				if self.printOut == True:
					print "Writing out: " + outMissing
				outFile.write("Part Number, Repair Number, Part Description, Alloc, De-Alloc, POS, RTW, Time Stamp\r")
				for entry in self.curDayUnbalanced:
					timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[7])
					outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + str(entry[3]) + "," + str(entry[4]) + "," + str(entry[5]) + "," + str(entry[6]) + "," + timeStampStr + "\r")
				outFile.close()

	def printReports(self, event='NONE'):
		self.writeCount()
		self.writeValid()
		self.writeMissing_Balanced()
		self.writeMissing_Unbalanced()
		self.writeMissing_POSdUnbalanced()
		#tkMessageBox.showinfo("Reports", "Reports have been saved to the folder on your desktop.")
		os.system("open " + self.dataFile[0:self.dataFile.rfind('/')+1])

	def __init__(self, inFile, dateMove, console=True):
		self.saleMovements = [ 947, 949, 951, 953, 955, 957, 959, 961, 975 ]
		self.reversalMovements = [ 948, 950, 952, 954, 956, 958, 960, 962, 976 ]
		self.dataFile = inFile
		self.dateFilter = dateMove
		self.title = []
		self.data = []
		self.unbalanced = []
		self.validParts = []
		self.curDayPOSdUnbalanced = []
		self.curDayUnbalanced = []
		self.otherDaysUnbalanced = []
		self.otherDays = []
		self.timeStamp = time.strftime("%m/%d/%y 12:00:00 AM")
		self.timeStamp = time.strptime(self.timeStamp, "%m/%d/%y %I:%M:%S %p")
		self.printOut = console
		self.ready=False
		self.box = [[]]

		try:
			self.readFile()
		except:
			tkMessageBox.showerror("Invalid File Format", "Reading in info failed.\nPlease select a Transaction History report\nof 4 full days with the default layout.")

#### Part Validation TKinter ####
def key(event):
	print "pressed", repr(event.char)


class batchScan:
	def checkBoxs(self, partNum, repairNum):
		boxNum = 0
		for box in self.PV.box:
			for entry in box:
				if partNum == entry[0] and repairNum == entry[1]:
					return boxNum + 1
			boxNum += 1
		return -1

	def scanIn(self, event='NONE'):
		partNum = self.input1.get()
		repairNum = self.input2.get()
		repairNum.upper()
		if partNum.count('-') != 1:
			tkMessageBox.showerror("Invalid Input", "Invalid formatted input. Please try again.")
		else:
			tempPartNum = partNum.split('-')
			if (len(tempPartNum) != 2) and (not tempPartNum[0].isdigit()) and (not tempPartNum[1].isdigit()):
				tkMessageBox.showerror("Invalid Input", "Invalid formatted input. Please try again.")
			elif repairNum[0] != 'R' or (not repairNum[1:].isdigit()):
				tkMessageBox.showerror("Invalid Input", "Invalid formatted input. Please try again.")
			else:
				result = self.checkBoxs(partNum, repairNum)
				if result != -1:
					tkMessageBox.showerror("Already Scanned", partNum + " - " + repairNum + "\nThis was already scanned into box " + str(result))
				else:
					print "Checking today..."
					result = self.PV.scanIn(self.PV.data, partNum, repairNum, self.boxNum)
					if result[0] >= 0:
						self.root.update_balanced()
						for box in self.root.boxs:
							box.update()
					elif result[0] == -2:
							tkMessageBox.showinfo("Consumable Part", "Part is a consumable, balanced, go ahead and scrap it.")
					else:
						print "Checking other days..."
						result = self.PV.scanIn(self.PV.otherDays, partNum, repairNum, self.boxNum)
						if result[0] >= 0:
							self.root.update_balanced()
							for box in self.root.boxs:
								box.update()
						elif result[0] == -2:
							tkMessageBox.showinfo("Consumable Part", "Part is a consumable, balanced, from a previous day, go ahead and scrap it.")
							self.root.update_balanced()
							for box in self.root.boxs:
								box.update()
						else:
							print "Checking today POS'd but unbalanced days..."
							why = ''
							found = False
							for line in self.root.PV.curDayPOSdUnbalanced:
								if repairNum == line[1]:
									if partNum == line[0]:
										if "LOANER" not in line[2]:
											why += "Found " + line[0] + " - " + line[1] + ",\n but it is unbalanced in movements."
											why += "\n344s: " + str(line[3])
											why += "\n343s: " + str(line[4])
											why += "\nSales: " + str(line[5])
											why += "\nRTWs: " + str(line[6])
											why += "\nManually run transaction history for this repair and investigate."
											found = True
											break
							if found == True:
								tkMessageBox.showerror("Invalid Scan", why)
							else:
								print "Checking today but unbalanced..."
								why = ''
								found = False
								for line in self.root.PV.curDayUnbalanced:
									if repairNum == line[1]:
										if partNum == line[0]:
											if "LOANER" not in line[2]:
												why += "Found " + line[0] + " - " + line[1] + ",\n but it is unbalanced in movements."
												why += "\n344s: " + str(line[3])
												why += "\n343s: " + str(line[4])
												why += "\nSales: " + str(line[5])
												why += "\nRTWs: " + str(line[6])
												why += "\nManually run transaction history for this repair and investigate."
												found = True
												break
								if found == True:
									tkMessageBox.showerror("Invalid Scan", why)
								else:
									tkMessageBox.showerror("Invalid Scan", "No match.\nSuggest running a manual transaction history.")
		self.input1.delete(0, END)
		self.input2.delete(0, END)
		self.input1.focus_set()

	def input2_focus(self, event='NONE'):
		self.input2.focus_set()

	def close(self):
		self.text1.destroy()
		self.text2.destroy()
		self.input1.destroy()
		self.input2.destroy()
		self.okBtn.destroy()

	def __init__(self, root, frame, PV, boxNum):
		self.root = root
		self.PV = PV
		self.boxNum = boxNum
		self.boxFrame = frame

		self.text1 = Label(self.boxFrame, text="Part Number:")
		self.text2 = Label(self.boxFrame, text="Repair Number:")
		self.input1 = Entry(self.boxFrame)
		self.input2 = Entry(self.boxFrame)
		self.okBtn = Button(self.boxFrame, text="Check", command=self.scanIn)
		self.text1.grid(row=2)
		self.text2.grid(row=3)
		self.input1.focus_set()
		self.input1.bind("<Return>", self.input2_focus)
		self.input1.grid(row=2, column=1)
		self.input2.grid(row=3, column=1)
		self.input2.bind("<Return>", self.scanIn)
		self.okBtn.grid(row=3, column=2)

class box:
	def addCaller(self, event='NONE'):
		self.root.addItem_balanced('NONE', self.boxNum)
		self.update()

	def removeCaller(self, event='NONE'):
		selected = self.box_list.curselection()
		selectedItems = []
		for entry in range(0, len(selected)):
			selectedItems.append(int(selected[entry]))
			print "Adding " + selected[entry]
		selectedItems = reversed(selectedItems)
		for index in selectedItems:
			print "Removing " + self.root.PV.box[self.boxNum][index][0] + ' - ' + self.root.PV.box[self.boxNum][index][1] + " from box " + str(self.boxNum)
			self.root.PV.data.append(self.root.PV.box[self.boxNum][index])
			self.root.PV.box[self.boxNum].pop(index)
		self.root.PV.data.sort()
		self.root.PV.box[self.boxNum].sort()
		self.root.update_balanced()
		for box in self.root.boxs:
			box.update()

	def update(self, event='NONE'):
		self.root.PV.box[self.boxNum].sort()
		calcHeight = 1
		if len(self.root.boxs) == 0:
			calcHeight = 30
		elif len(self.root.boxs) >= 1:
			calcHeight = (30 / (len(self.root.boxs)))-1
		self.box_list.config(height=calcHeight)
		self.box_list.delete(0, END)
		if len(self.root.PV.box[self.boxNum]) > 0:
			for line in self.root.PV.box[self.boxNum]:
				line_str = ' ' + str(line[0]).ljust(12) + str(line[1]).ljust(15)
				self.box_list.insert(END, line_str + "\n")
			for entry in range(1, len(self.root.PV.box[self.boxNum]), 2):
				self.box_list.itemconfigure(entry, background="pink")

	def deleteBox(self, event='NONE'):
		if len(self.root.PV.box[self.boxNum]) > 0:
			result = tkMessageBox.askokcancel("Delete Box", "Box " + str(self.boxNum+1) + " has items added to it,\n if you continue the items will be removed and will need to be processed again.")
			if result:
				for line in self.root.PV.box[self.boxNum]:
					self.root.PV.data.append(line)
				self.root.PV.data.sort()
				self.root.PV.box.pop(self.boxNum)
				self.root.removeBox('NONE', self.boxNum)
				self.batchScanHandler.close()
				self.boxFrame.destroy()
		else:
			self.root.PV.box.pop(self.boxNum)
			self.buttonFrame.destroy()
			self.box_list.destroy()
			self.root.removeBox('NONE', self.boxNum)
			self.batchScanHandler.close()
			self.boxFrame.destroy()

	def __init__(self, root, boxsFrame, boxNum):
		self.boxsFrame = boxsFrame
		self.root = root
		self.boxNum = boxNum

		self.boxFrame = Frame(self.boxsFrame, background="light grey", padx=2, pady=2)
		self.boxFrame.grid(row=self.boxNum)

		self.buttonFrame = Frame(self.boxFrame, background="light grey")
		self.buttonFrame.grid(row=0, column=0)

		self.box_button1 = Tkinter.Button(self.buttonFrame, text=">", command=self.addCaller)
		self.box_button1.grid(row=0, column=0)

		self.box_button2 = Tkinter.Button(self.buttonFrame, text="<", command=self.removeCaller)
		self.box_button2.grid(row=1, column=0)

		self.batchScanHandler = batchScan(self.root, self.boxFrame, self.root.PV, self.boxNum)

		if self.boxNum > 0:
			self.box_button4 = Tkinter.Button(self.buttonFrame, text="Remove Box", command=self.deleteBox)
			self.box_button4.grid(row=2, column=0)

		self.box_list = Tkinter.Listbox(self.boxFrame, height=30, width=28, selectmode=EXTENDED, font=("Consolas", '14'))
		self.update()
		self.box_list.grid(row=0, column=1, padx=10, pady=10)
		self.box_list.focus_set()
		self.box_list.selection_set(0)

class getData:

	def selectAll_balanced(self, event='NONE'):
		self.balance.dataText_balanced.selection_set(0,END)

	def update_balanced(self, event='NONE'):
		self.PV.data.sort()
		self.balance.dataText_balanced.delete(0,END)
		for line in self.PV.data:
			line_str = ' ' + str(line[0]).ljust(12) + str(line[1]).ljust(15) + str(line[2]).ljust(42) + str(time.strftime("%m/%d/%y %I:%M:%S %p", line[7])).rjust(10)
			self.balance.dataText_balanced.insert(END, line_str + "\n")
		for slot in range(0, len(self.PV.data)):
			self.balance.dataText_balanced.itemconfigure(slot, background="white")
		for slot in range(1, len(self.PV.data), 2):
			self.balance.dataText_balanced.itemconfigure(slot, background="pink")

	def addItem_balanced(self, event='NONE', boxNum=-10):
		if boxNum >= 0:
			items = map(int, self.balance.dataText_balanced.curselection())
			if items:
				for entry in range(len(items)-1, -1, -1):
					result = self.PV.scanIn(self.PV.data, self.PV.data[int(items[entry])][0], self.PV.data[int(items[entry])][1], boxNum)
					if result != -1:
						self.balance.dataText_balanced.delete(items[entry])
						self.update_balanced()
						for box in self.boxs:
							box.update()
				position = items[0]
				while True:
					if position > len(self.PV.data)-1:
						position -= 1
					else:
						break
				self.balance.dataText_balanced.activate(position)
				self.balance.dataText_balanced.selection_set(position)

	def addItem_balanced_prompt_return(self, event='NONE', boxNum=-10):
		self.prompt.destroy()
		self.addItem_balanced(event, boxNum)

	def addItem_balanced_prompt(self, event='NONE'):
		self.prompt = Toplevel()
		self.prompt.title("Box Input")
		frame = Frame(self.prompt)
		label = Label(frame, text="Enter in box number:")
		entry = Entry(frame)
		btn = Button(frame, text="Done", command=lambda: self.addItem_balanced_prompt_return("event", int(entry.get())-1 ))
		entry.focus_set()
		entry.bind("<Return>", lambda event: self.addItem_balanced_prompt_return(event, int(entry.get())-1) )
		##entry.bind("Enter", lambda event: self.addItem_balanced_prompt_return(event, int(entry.get())-1) )
		frame.grid(row=0)
		label.grid(row=0)
		entry.grid(row=1)
		btn.grid(row=2)

	def getData(self, event='NONE', ):
		self.PV = PartValidation(self.userFilePath, self.dateFilter, True)
		self.printData(self)

	def printData_balanced(self, event='NONE'):
		self.PV.data.sort()
		## POS'd and balanced listbox
		self.balance = LabelFrame(self.mainFrame, text="POS'd and Balanced:", background='light grey')
		self.balance.dataText_balanced = Tkinter.Listbox(self.balance, height=35, width=91, selectmode=EXTENDED, font=("Consolas", '14'))
		self.balance.dataText_balanced.bind("<Control-a>", self.selectAll_balanced)
		self.balance.dataText_balanced.bind("<Return>", self.addItem_balanced_prompt)
		for line in self.PV.data:
			line_str = ' ' + str(line[0]).ljust(12) + str(line[1]).ljust(15) + str(line[2]).ljust(42) + str(time.strftime("%m/%d/%y %I:%M:%S %p", line[7])).rjust(10)
			self.balance.dataText_balanced.insert(END, line_str + "\n")
		for entry in range(1, len(self.PV.data), 2):
			self.balance.dataText_balanced.itemconfigure(entry, background="pink")
		self.balance.dataText_balanced.grid(columnspan=2, padx=10, pady=10)
		self.balance.dataText_balanced.focus_set()
		self.balance.dataText_balanced.selection_set(0)

	def printData_unbalanced(self, event='none'):
		## POS'd but unbalanced listbox
		self.unbalance = LabelFrame(self.mainFrame, text="Moved today:", background='light grey')

		title = Label(self.unbalance, text=' ' + str("Part Number").ljust(15) + str("Repair Number").ljust(20) + str("Part Description").ljust(72) + str("Alloc Dealloc POS RTW").ljust(35) + str("Date").ljust(72))
		title.grid(row=2, sticky=W, padx=10)

		self.PV.curDayUnbalanced.sort()

		if len(self.PV.curDayUnbalanced) > 15:
			todayUnbalanced_height = 15
		else :
			todayUnbalanced_height = len(self.PV.curDayUnbalanced)

		self.unbalance.dataText_unbalanced = Tkinter.Listbox(self.unbalance, height=todayUnbalanced_height, width=131, selectmode=EXTENDED, font=("Consolas", '14'))
		for line in self.PV.curDayUnbalanced:
			if "LOANER" not in line[2]:
				why = str(line[3]).ljust(5) + str(line[4]).ljust(5) + str(line[5]).ljust(5) + str(line[6]).ljust(5)
				line_str = ' ' + str(line[0]).ljust(12) + str(line[1]).ljust(15) + str(line[2]).ljust(42) + why.ljust(22) + str(time.strftime("%m/%d/%y %I:%M:%S %p", line[7])).rjust(10)
				self.unbalance.dataText_unbalanced.insert(END, line_str + "\n")
		for entry in range(1, len(self.PV.curDayUnbalanced), 2):
			self.unbalance.dataText_unbalanced.itemconfigure(entry, background="pink")
		self.unbalance.dataText_unbalanced.grid(row=3,padx=10, pady=10)

	def printData_POSdUnbalanced(self, event='none'):
		## POS'd but unbalanced listbox
		self.POSdUnbalance = LabelFrame(self.mainFrame, text="POS'd But Unbalanced:", background='light grey')

		title = Label(self.POSdUnbalance, text=' ' + str("Part Number").ljust(15) + str("Repair Number").ljust(20) + str("Part Description").ljust(72) + str("Alloc Dealloc POS RTW").ljust(35) + str("Date").ljust(72))
		title.grid(row=0, sticky=W, padx=10)

		self.PV.curDayPOSdUnbalanced.sort()

		if len(self.PV.curDayPOSdUnbalanced) > 15:
			todayUnbalanced_height = 15
		else :
			todayUnbalanced_height = len(self.PV.curDayPOSdUnbalanced)

		self.POSdUnbalance.dataText_POSdUnbalanced = Tkinter.Listbox(self.POSdUnbalance, height=todayUnbalanced_height, width=131, selectmode=EXTENDED, font=("Consolas", '14'))
		for line in self.PV.curDayPOSdUnbalanced:
			if "LOANER" not in line[2]:
				why = str(line[3]).ljust(5) + str(line[4]).ljust(5) + str(line[5]).ljust(5) + str(line[6]).ljust(5)
				line_str = ' ' + str(line[0]).ljust(12) + str(line[1]).ljust(15) + str(line[2]).ljust(42) + why.ljust(22) + str(time.strftime("%m/%d/%y %I:%M:%S %p", line[7])).rjust(10)
				self.POSdUnbalance.dataText_POSdUnbalanced.insert(END, line_str + "\n")
		for entry in range(1, len(self.PV.curDayPOSdUnbalanced), 2):
			self.POSdUnbalance.dataText_POSdUnbalanced.itemconfigure(entry, background="pink")
		self.POSdUnbalance.dataText_POSdUnbalanced.grid(row=1,padx=10, pady=10)

	def addBox(self, event='NONE'):
		self.PV.addBox()
		self.boxs.append(box(self, self.boxsFrame, len(self.boxs)))
		for index in range(0, len(self.boxs)):
			self.boxs[index].update()

	def removeBox(self, event='NONE', boxNum=-10):
		if boxNum >= 0:
			self.boxs.pop(boxNum)
		for index in range(0, len(self.boxs)):
			self.boxs[index].boxNum = index
			self.boxs[index].update()

	def printData(self, event="NONE"):
		## Buttons row

		self.btnFrame = Frame(self.mainFrame)
		self.btnFrame.grid(row=0, columnspan=3)

		self.btnFrame.addToBox = Tkinter.Button(self.btnFrame, text="Add to Box", command=lambda: self.addItem_balanced_prompt() )
		self.btnFrame.addToBox.grid(row=0, column=0)

		self.btnFrame.dataText_balanced_reportsBtn = Button(self.btnFrame, text="Save Reports", command=self.PV.printReports, background="light grey")
		self.btnFrame.dataText_balanced_reportsBtn.grid(row=0, column=2)

		## Balanced Data box
		self.printData_balanced()
		self.balance.grid(row=1, column=0, columnspan=3, sticky=W)

		## Unbalanced Data box
		self.printData_POSdUnbalanced()
		self.POSdUnbalance.grid(row=2, column=0, columnspan=5, sticky=W)

		## Unbalanced Data box
		self.printData_unbalanced()
		self.unbalance.grid(row=3, column=0, columnspan=5, sticky=W)


		## SRTW boxes
		self.boxsFrame = LabelFrame(self.mainFrame, text="Boxes:", background="light grey")
		self.boxs = []
		self.boxs.append(box(self, self.boxsFrame, 0))
		self.boxsFrame.grid(row=1, column=3, columnspan=2)

		self.boxsAdd = Button(self.mainFrame, text="Add Box", command=lambda: self.addBox())
		self.boxsAdd.grid(row=0, column=4, sticky=W)

	def __init__(self, mainFrame, userFilePath, dateFilter, timeFilter):
		self.mainFrame = mainFrame
		self.userFilePath = userFilePath
		self.dateFilter = dateFilter
		self.timeFilter = timeFilter

		#### Fix addItem to see if the current selection is in balanced, unbalanced, or both...
		self.getData()

class mainWindow:
	def getFile(self, event='NONE'):
		self.userFilePath = str(tkFileDialog.askopenfilename(filetypes=[('Excel','.xlsx')]))
		self.userFile_entry.insert(0, self.userFilePath)
		self.dateFrame.focus_set()

	def moveFocus(self, event='NONE'):
		self.time_hour_option.focus_set()

	def getDataAndStart(self, event='NONE'):
		if os.path.exists(self.userFilePath):
			self.dateFilter = time.strptime(self.date_month.get() + "/" + self.date_day.get() + "/" + self.date_year.get() + " " + str(self.time_hour.get()) + ":" + str(self.time_min.get()) + " " + str(self.time_AmPm.get()), "%m/%d/%Y %I:%M %p")
			self.userFile_label.destroy()
			self.userFile_button.destroy()
			self.userFile_entry.destroy()
			self.dateFrame.destroy()
			self.dateFilter_Days_label.destroy()
			self.dateFilter_Done_button.destroy()
			self.timeFrame.destroy()
			self.dateFilter_Time_label.destroy()
			self.getDataHandler = getData(self.mainFrame, self.userFilePath, self.dateFilter, True)
		else:
			tkMessageBox.showerror("Invalid File", "Please select a valid file.")

	def updateMonth(self, arg1='None', arg2='None', arg3='None'):
		top = int(self.date_month.get())
		if top == 2:
			if int(self.date_year.get()) % 4 == 0:
				days = range(1, 30)
			else:
				days = range(1, 29)
		elif top in [ 1, 3, 5, 7, 8, 10, 12]:
			days = range(1, 32)
		else:
			days = range(1, 31)

		for index in range(0, len(days)):
			days[index] = str(days[index]).zfill(2)
		try:
			self.date_day_option.destroy()
		except:
			pass
		self.date_day_option = apply(OptionMenu, (self.dateFrame, self.date_day) + tuple(days))
		self.date_day_option.grid(row=0, column=1, sticky=W)

	def __init__(self, userFile=''):
		#baseDir = sys.argv[0][0:sys.argv[0].rfind('/')+1]

		self.userFilePath = ""

		if len(sys.argv) > 1:
			userFile = sys.argv[len(sys.argv)-1]
			self.userFilePath = userFile

		root = Tkinter.Tk()
		root.resizable(0,0)

		root.title("Part Validation v2.0")
		self.mainFrame = Tkinter.Frame(root, padx=10, pady=10, background="light grey")
		self.mainFrame.grid()

		self.userFile_label = Label(self.mainFrame, text="Select the XLSX File:", background="light grey")
		self.userFile_entry = Entry(self.mainFrame)
		self.userFile_entry.insert(0, userFile)
		self.userFile_button = Button(self.mainFrame, text="Open...", command=self.getFile, background="light grey")
		self.userFile_button.bind("<Return>", self.getFile)

		self.dateFilter_Days_label = Label(self.mainFrame, text="How many day's worth of SRTW (Today = 1):", background="light grey")
		self.dateFrame = Frame(self.mainFrame)

		self.curDate = time.strftime("%m/%d/%y")
		self.curDate = time.strptime(self.curDate, "%m/%d/%y")

		self.date_day = StringVar(root)
		self.date_day.set(str(self.curDate[2]).zfill(2))

		self.date_month = StringVar(root)
		self.date_month.set(str(self.curDate[1]).zfill(2))

		self.date_year = StringVar(root)
		self.date_year.set(str(self.curDate[0]))

		## Month Selector
		self.date_month.trace("w", self.updateMonth)
		self.date_month_option = OptionMenu(self.dateFrame, self.date_month, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
		self.date_month_option.grid(row=0, column=0, sticky=W)

		## Day Selector
		self.updateMonth()

		## Year Selector
		self.date_year.trace("w", self.updateMonth)
		years = range(self.curDate[0]-5, self.curDate[0]+3)
		self.date_year_option = apply(OptionMenu, (self.dateFrame, self.date_year) + tuple(years))
		self.date_year_option.grid(row=0, column=2, sticky=W)

		self.dateFrame.bind("<Return>", self.moveFocus)
		self.dateFrame.grid(row=1, column=1, columnspan=2, sticky=W)

		self.dateFilter_Time_label = Label(self.mainFrame, text="From what time today do you wish to check\n (Example: 11:30 am) (12:00 AM = all day):", background="light grey")

		self.timeFrame = Frame(self.mainFrame)
		self.timeFrame.grid(row=2, column=1, columnspan=2)

		## Hour selector
		self.time_hour = StringVar(root)
		self.time_hour.set("12")
		self.time_hour_option = OptionMenu(self.timeFrame, self.time_hour, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
		self.time_hour_option.grid(row=0, column=0, sticky=W)

		## Minute Selector
		self.time_min = StringVar(root)
		self.time_min.set("00")
		minutes = range(0, 60)
		for index in range(0, 60):
			minutes[index] = str(minutes[index]).zfill(2)
		self.time_min_option = apply(OptionMenu, (self.timeFrame, self.time_min) + tuple(minutes))
		self.time_min_option.grid(row=0, column=1, sticky=W)

		## AM / PM selector
		self.time_AmPm = StringVar(root)
		self.time_AmPm.set("AM")

		self.time_AmPm_option = OptionMenu(self.timeFrame, self.time_AmPm, "AM", "PM")
		self.time_AmPm_option.grid(row=0, column=2, sticky=W)


		self.dateFilter_Done_button = Button(self.mainFrame, text="Done", command=self.getDataAndStart, background="light grey")

		self.userFile_label.grid(row=0, sticky=E)
		self.userFile_button.grid(row=0, column=1, sticky=W)
		self.userFile_button.focus_set()
		self.userFile_entry.grid(row=0, column=2, sticky=W)

		self.dateFilter_Days_label.grid(row=1, sticky=E)

		self.dateFilter_Time_label.grid(row=2, sticky=E)

		self.dateFilter_Done_button.grid(row=3, column=3, sticky=E)

		#mainFrame.getDataBtn = Tkinter.Button(mainFrame, text="Open...", command=getDataHandler.getUserData, width=7, height=2)
		#mainFrame.getDataBtn.grid(row=2)

		root.mainloop()

if __name__ == "__main__":
	sys.path.append("/Users/apple/Documents/Part Validation/")
	import xlrd
	root = mainWindow()