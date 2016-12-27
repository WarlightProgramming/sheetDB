##############
# worksheet.py
# handles single worksheet
##############

from credentials import Credentials 
from functions import *
from errors import *
import gspread

# main class
class Worksheet(object):

    def __init__(self, worksheet):
        self.sheet = worksheet
        self.client = worksheet.client

    ## built-in methods

    @updateClient
    def __eq__(self, other):
        return (self.sheet.id == other.sheet.id)

    @updateClient
    def __hash__(self):
        return hash(self.sheet.id)

    ## read operations

    @updateClient
    def getColCount(self):
        return self.sheet.col_count

    @updateClient
    def getRowCount(self):
        return self.sheet.row_count

    @updateClient
    def getTitle(self):
        return self.sheet.title

    @updateClient
    def getID(self):
        return self.sheet.id

    @updateClient
    def getUpdateTime(self):
        return self.sheet.updated

    @updateClient
    def getClient(self):
        return self.client

    @updateClient
    def getCol(self, col):
        if (col <= self.getColCount()):
            return self.sheet.col_values(col)
        else:
            raise SheetError("Nonexistent column!")

    @updateClient
    def getRow(self, row):
        if (row <= self.getRowCount()):
            return self.sheet.row_values(row)
        else:
            raise SheetError("Nonexistent row!")

    @updateClient
    def export(self, format="csv"):
        return self.sheet.export(format)

    @updateClient
    def exportCSV(self):
        return self.export("csv")

    @updateClient
    def findall(self, query):
        return self.sheet.findall(query)

    @updateClient
    def getRecords(self, zero=False, head=1):
        return self.sheet.get_all_records(empty2zero=zero,
                                          head=head)

    @updateClient
    def getLabel(self, row, col):
        return self.sheet.get_addr_int(row, col)

    @updateClient
    def getRowCol(self, label):
        return self.sheet.get_int_addr(label)

    @updateClient
    def getAll(self):
        return self.sheet.get_all_values()

    @updateClient
    def getLabeledRange(self, rangeLabel):
        return self.sheet.range(rangeLabel)

    @updateClient
    def __getRangeLabel(self, row1, col1, row2, col2):
        return (self.getLabel(row1, row2) + ":" +
                self.getLabel(row2, col2))

    @updateClient
    def getNumberedRange(self, row1, col1, row2, col2):
        return self.getLabeledRange(self.getRangeLabel(row1,
                                    col1, row2, col2))

    @updateClient
    def getRange(self, rangeLabel):
        if (type(rangeLabel) == str):
            return self.getLabeledRange(self, rangeLabel)
        elif (type(rangeLabel) == list or 
              type(rangeLabel) == tuple):
            if (len(rangeLabel) == 2):
                row1 = rangeLabel[0][0]
                col1 = rangeLabel[0][1]
                row2 = rangeLabel[1][0]
                col2 = rangeLabel[1][1]
            elif (len(rangeLabel) == 4):
                row1 = rangeLabel[0]
                col1 = rangeLabel[1]
                row2 = rangeLabel[2]
                col2 = rangeLabel[3]
            return self.getNumberedRange(row1, col1, row2, col2)

    @updateClient
    def fetchCell(self, loc):
        if (type(loc) == tuple):
            return self.sheet.cell(loc[0], loc[1])
        elif (type(loc) == str):
            return self.sheet.acell(loc)

    @updateClient
    def fetchCellValue(self, loc):
        cell = self.fetchCell(loc)
        return cell.value

    @updateClient
    def fetchCellInputValue(self, loc):
        cell = self.fetchCell(loc)
        return cell.input_value

    ## write operations

    @updateClient
    def addCols(self, count):
        self.sheet.add_cols(count)

    @updateClient
    def addRows(self, count):
        self.sheet.add_rows(count)

    @updateClient
    def appendRows(self, *rows):
        for row in rows:
            self.sheet.append_row(row)

    @updateClient
    def insertRow(self, values, index=1):
        self.sheet.insert_row(values, index)

    @updateClient
    def fillRow(self, rowNum, newValues):
        if (len(newValues) > self.getColCount()):
            self.addCols(len(newValues) - colCount)
        for x in xrange(len(newValues)):
            self.updateCell((rowNum, x), newValues[x])

    @updateClient
    def fillCol(self, colNum, newValues):
        if (len(newValues) > self.getRowCount()):
            self.addRows(len(newValues) - rowCount)
        for x in xrange(len(newValues)):
            self.updateCell((x, colNum), newValues[x])

    @updateClient
    def appendCols(self, *cols):
        rowCount = self.getRowCount()
        colNum = self.getColCount()
        self.addCols(len(cols))
        for col in cols:
            colNum += 1
            self.fillCol(colNum, col)

    @updateClient
    def updateCell(self, loc, value):
        if (type(loc) == tuple):
            self.sheet.update_cell(loc[0], loc[1], value)
        elif (type(loc) == str):
            self.sheet.update_acell(loc, value)

    @updateClient
    def resize(self, rowCount, colCount):
        self.sheet.resize(rowCount, colCount)

    @updateClient
    def deleteSheet(self):
        self.del_worksheet(self.sheet)

    @updateClient
    def delete(self):
        self.deleteSheet()