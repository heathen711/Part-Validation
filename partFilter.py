import os
import sys
import csv
import subprocess
import xlrd
import time
from operator import itemgetter

class partFilter:

    def readInXLSX(self):
        print "Opening Spreedsheet... (may take a while)"
        error = False
        try:
            workbook = xlrd.open_workbook(self.dataFile)
        except IOError:
            print "Could not open file! Check path and please close Excel."
            ## Set Control to fail
            error = True
        if (error == False):
            worksheet = workbook.sheet_by_name('Sheet1')
            num_rows = worksheet.nrows - 1
            num_cells = worksheet.ncols - 1
            curr_row = 0
            gotTitle = False
            while curr_row < num_rows:
                curr_row += 1
                row = worksheet.row(curr_row)
                curr_cell = -1
                exclude = [0, 4, 7, 11, 13, 14, 15]
                row_values = []
                while curr_cell < num_cells:
                    curr_cell += 1
                    # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
                    if (curr_cell not in exclude):
                        cell_type = worksheet.cell_type(curr_row, curr_cell)
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
                row_values[2] = row_values[2].replace('"', '')
                if ((int(row_values[3]) in [ 251, 252, 344, 343 ]) or (int(row_values[3]) in self.saleMovements) or (int(row_values[3]) in self.reversalMovements) or (int(row_values[3]) in self.bohMovements)):
                    if (("922-" in row_values[1]) or ("923-" in row_values[1]) or ("076-" in row_values[1]) or ("661-" in row_values[1])):
                        ## Add the info to the data array
                        if row_values[5] != '':
                            location = row_values[5]
                        else:
                            location = row_values[6]
                        # [ Repair Number, Part Number, Part Description, Movement, Stock Type, TimeStamp, Location Code ]
                        self.rawData.append([ row_values[7], row_values[1], row_values[2], row_values[3], row_values[4], row_values[0], location ])
        self.rawData.sort()
        print "Read in all data."
        return error

    def addMovement(self, movement):
        print "Adding movement..."
        found = False

        print movement[3]
        if int(movement[3]) in self.bohMovements:
            self.bohAdjustments.append([ movement[1], movement[0], movement[2], 0, 0, 0, 0, movement[5]])
            print "Found a boh adjustment..."
        else:
            ## self.allocationTable = [ repairNum, partNum, partDescription, # of 344, # of 343, # of sold (range(947,961) & 251), # of returns 252, timeStamp ]
            for row in self.allocationTable:
                if ((movement[0] == row[1]) and (movement[1] == row[0])):
                    if (movement[3] == "344"):
                        row[3] += 1
                        row[7] = movement[5]
                    elif (movement[3] == "343"):
                        row[4] += 1
                        row[7] = movement[5]
                    elif (int(movement[3]) in self.reversalMovements):
                        row[5] -= 1
                        row[7] = movement[5]
                    elif (int(movement[3]) in self.saleMovements):
                        row[5] += 1
                    elif (movement[3] == "251" and movement[6] == "001"):
                        row[5] += 1
                        row[7] = movement[5]
                    elif (movement[3] == "251" and movement[6] == "002"):
                        row[6] -= 1
                        row[7] = movement[5]
                    elif (movement[3] == "252" and movement[6] == "001"):
                        row[5] -= 1
                        row[7] = movement[5]
                    elif (movement[3] == "252" and movement[6] == "002"):
                        row[6] += 1
                        row[7] = movement[5]
                    found = True

            if found == False:
                self.allocationTable.append([ movement[1], movement[0], movement[2], 0, 0, 0, 0, ""])
                index = len(self.allocationTable)-1
                if (movement[3] == "344"):
                    self.allocationTable[index][3] += 1
                    self.allocationTable[index][7] = movement[5]
                elif (movement[3] == "343"):
                    self.allocationTable[index][4] += 1
                    self.allocationTable[index][7] = movement[5]
                elif (int(movement[3]) in self.reversalMovements):
                    self.allocationTable[index][5] -= 1
                    self.allocationTable[index][7] = movement[5]
                elif (int(movement[3]) in self.saleMovements):
                    self.allocationTable[index][5] += 1
                    self.allocationTable[index][7] = movement[5]
                elif (movement[3] == "251" and movement[6] == "001"):
                    self.allocationTable[index][5] += 1
                    self.allocationTable[index][7] = movement[5]
                elif (movement[3] == "251" and movement[6] == "002"):
                    self.allocationTable[index][6] -= 1
                    self.allocationTable[index][7] = movement[5]
                elif (movement[3] == "252" and movement[6] == "001"):
                    self.allocationTable[index][5] -= 1
                    self.allocationTable[index][7] = movement[5]
                elif (movement[3] == "252" and movement[6] == "002"):
                    self.allocationTable[index][6] += 1
                    self.allocationTable[index][7] = movement[5]
            if self.allocationTable[0][0] == '':
                self.allocationTable.pop(0)
            self.allocationTable.sort()

    def allocateFilter(self):
        print "Building Movement Filter..."
        ## rawData = [ repairNum, partNum, partDescription, movementCode, stockType, repairNum, timeStamp ]
        ## self.allocationTable = [ repairNum, partNum, partDescription, # of 344, # of 343, # of sold (range(947,961) & 251), # of returns 252, timeStamp ]

        ## look for repairs returned at POS
        self.returnedAtPOS = []
        for index in range(0, len(self.rawData)):
            if (self.rawData[index][3] == "252" and self.rawData[index][4] == "Block"):
                    self.returnedAtPOS.append(self.rawData[index])
                    print "Found repair returned by manager."

        ## Build array of movements for repair numbers and part numbers
        self.allocationTable.append([ "", "", "", 0, 0, 0, 0, ""])
        for movement in self.rawData:
            if (int(movement[3]) in self.bohMovements):
                temp = movement
                temp[0] = movement[3]
                self.bohAdjustments.append(temp)
            else :
                self.addMovement(movement)

    def unbalanced(self):
        print "Building Unbalanced Filter..."
        unbalanced = []
        for index in range(0, len(self.allocationTable)):
            added = False
            allocated = self.allocationTable[index][3] - self.allocationTable[index][4]
            if (allocated != self.allocationTable[index][5]):
                unbalanced.append(self.allocationTable[index])
                added = True
            if ((self.allocationTable[index][5] != self.allocationTable[index][6]) and (added == False)):
                unbalanced.append(self.allocationTable[index])
        unbalanced.sort()

        outUnbalanced = self.dataFile[:self.dataFile.rfind('.')] + "_unbalanced.jay.csv"
        opened = False
        try:
            ## Open valid jay in write mode
            outFile = open(outUnbalanced, "w")
            opened = True
        except IOError:
            print "Could not open valid output file!"

        if (opened == True):
            print "Writing out: " + outUnbalanced
            outFile.write("Part Number,Repair Number,Part Description,344,343,Sale,RTW,Time,\r")
            for entry in unbalanced:
                timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[7])
                outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + str(entry[3]) + "," + str(entry[4]) + "," + str(entry[5]) + "," + str(entry[6]) + "," + timeStampStr + "\r")
            if (len(self.returnedAtPOS) > 0):
                outFile.write(",Returned:,\r")
                for entry in self.returnedAtPOS:
                    timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[5])
                    outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + " , , , , " + timeStampStr + "\r")
            if (len(self.bohAdjustments) > 0):
                outFile.write(",Boh Adjusted:,\r")
                for entry in self.bohAdjustments:
                    timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[5])
                    outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + " , , , , " + timeStampStr + "\r")
            outFile.close()

    def full(self):
        print "Building Full List..."
        outUnbalanced = self.dataFile[:self.dataFile.rfind('.')] + "_full.jay.csv"
        opened = False
        try:
            ## Open valid jay in write mode
            outFile = open(outUnbalanced, "w")
            opened = True
        except IOError:
            print "Could not open valid output file!"

        if (opened == True):
            print "Writing out: " + outUnbalanced
            outFile.write("Part Number,Repair Number,Part Description,344,343,Sale,RTW,Time,\r")
            for entry in self.allocationTable:
                timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[7])
                outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + str(entry[3]) + "," + str(entry[4]) + "," + str(entry[5]) + "," + str(entry[6]) + "," + timeStampStr + "\r")
            if (len(self.returnedAtPOS) > 0):
                outFile.write(",Returned:,\r")
                for entry in self.returnedAtPOS:
                    timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[5])
                    outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + " , , , , " + timeStampStr + "\r")
            if (len(self.bohAdjustments) > 0):
                outFile.write(",Boh Adjusted:,\r")
                for entry in self.bohAdjustments:
                    timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[5])
                    outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + " , , , , " + timeStampStr + "\r")
            outFile.close()

    def POS(self):
        print "Building Sold Filter..."
        POSd = []
        for index in range(0, len(self.allocationTable)):
            if (self.allocationTable[index][5] > 0):
                added = False
                allocated = self.allocationTable[index][3] - self.allocationTable[index][4]
                if (allocated == self.allocationTable[index][5]):
                    POSd.append(self.allocationTable[index])
                    added = True
        POSd.sort()

        outUnbalanced = self.dataFile[:self.dataFile.rfind('.')] + "_POS'd.jay.csv"
        opened = False
        try:
            ## Open valid jay in write mode
            outFile = open(outUnbalanced, "w")
            opened = True
        except IOError:
            print "Could not open valid output file!"

        if (opened == True):
            print "Writing out: " + outUnbalanced
            outFile.write("Part Number,Repair Number,Part Description,344,343,Sale,RTW,Time,\r")
            for entry in POSd:
                timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[7])
                outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + str(entry[3]) + "," + str(entry[4]) + "," + str(entry[5]) + "," + str(entry[6]) + "," + timeStampStr + "\r")
            if (len(self.returnedAtPOS) > 0):
                outFile.write(",Returned:,\r")
                for entry in self.returnedAtPOS:
                    timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[5])
                    outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + " , , , , " + timeStampStr + "\r")
            if (len(self.bohAdjustments) > 0):
                outFile.write(",Boh Adjusted:,\r")
                for entry in self.bohAdjustments:
                    timeStampStr = time.strftime("%m/%d/%y %I:%M:%S %p", entry[5])
                    outFile.write(entry[0] + "," + entry[1] + ",\"" + entry[2] + "\"," + " , , , , " + timeStampStr + "\r")
            outFile.close()

    def __init__(self, whatToFilter, dataFile):
        print "Starting filter builder..."
        self.dataFile = dataFile
        self.saleMovements = [ 947, 949, 951, 953, 955, 957, 959, 961, 975 ]
        self.reversalMovements = [ 948, 950, 952, 954, 956, 958, 960, 962, 976 ]
        self.bohMovements = [ 941, 942, 701, 702 ]

        ## Format = [ Part Number, Description, Repair Number, Time (as time.struct_time) ]
        self.rawData = []
        self.allocationTable = []
        self.returnedAtPOS = []
        self.bohAdjustments = []

        self.timeStamp = time.strptime(time.strftime("%m/%d/%y 12:00:00 AM"), "%m/%d/%y %I:%M:%S %p")

        self.readInXLSX()
        self.allocateFilter()

        if whatToFilter == 1:
            self.full()
        elif whatToFilter == 2:
            self.unbalanced()
        elif whatToFilter == 3:
            self.POS()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[2].lower() == "full":
            root = partFilter(1, sys.argv[1])
        elif sys.argv[2].lower() == "unbalanced":
            root = partFilter(2, sys.argv[1])
        elif sys.argv[2].lower() == "pos":
            root = partFilter(3, sys.argv[1])
    else:
        print "################################"
        print "####     Part Validation    ####"
        print "################################"
        print "Use: Export SAP Transaction History of a single part number at least two weeks back, Drag and Drop, Enter, Open the export_POSd.jay.csv."

        dataFile = raw_input("Drag the XSLX file here and press enter: ")
        if (dataFile[len(dataFile)-1] == ' '):
            dataFile = dataFile[0:len(dataFile)-1]
        dataFile = dataFile.replace('\\', '')

        print "Enter in choice for filter:"
        print "1 - Full allocation log,"
        print "2 - Allocation issues,"
        print "3 - POS'd log,"
        print "4 - Exit"
        choice = raw_input("Choice: ")
        if choice.isdigit():
            choice = int(choice)
            if choice in range(1,5):
                root = partFilter(choice, dataFile)
