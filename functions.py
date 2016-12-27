# functions.py
## classless utility functions

from datetime import datetime
from errors import *
import string
import re
import copy

# function decorators

## updateClient
### updates client if expired
def updateClient(func):
    def func_wrapper(self, *args):
        if (self.client.auth.access_token_expired):
            self.client.login()
        return func(self, *args)
    return func_wrapper

# classless functions

## trim
### removes leading/trailing spaces/newlines or other
### characters in trimSet
def trim(string, trimSet=set([" ", "\n"])):
    while len(string) > 0 and string[0] in trimSet:
        string = string[1:]
    while len(string) > 0 and string[-1] in trimSet:
        string = string[:-1]
    return string

## getCellNumber
### given a cell label
### retrieves cell number
def getCellNumber(label):
    letters = ""
    numbers = ""
    for char in label:
        if char in string.digits:
            numbers += char
        elif char in string.ascii_letters:
            letters += char
    if (letters not in label or numbers not in label):
        raise DataError("Invalid label!")
    colLabel = letters
    rowNum = int(numbers)
    return (rowNum, getColumnNumber(colLabel))

## getCellLabel
### given a row number and a column number
### returns alphabetical cell label
def getCellLabel(rowNumber, colNumber):
    return getColumnLabel(colNumber) + str(rowNumber)

## getAlpha
### returns capital letter corresponding to number
### works for numbers 1 through 26
def getAlpha(number):
    return chr(number + 64)

## getNumber
### returns the number corresponding to a capital letter
def getNumber(alpha):
    return ord(alpha) - 64

## getColumnLabel
### given a column number (>= 1),
### returns column alphabetical label
def getColumnLabel(colNumber):
    colLabel = ""
    while (colNumber > 0):
        remainder = (colNumber - 1) / 26
        useNumber = colNumber - (26 * remainder)
        colLabel = getAlpha(useNumber) + colLabel
        colNumber = remainder
    return colLabel

## getColumnNumber
### given a column number (string),
### returns the column number
def getColumnNumber(colLabel):
    colNumber = 0
    while (len(colLabel) > 0):
        exponent = len(colLabel) - 1
        currentChar = colLabel[0]
        colLabel = colLabel[1:]
        colNumber += getNumber(currentChar) * (26 ** exponent)
    return colNumber

## getShiftedReference
### given an original label (string),
### and a shift vector (tuple or list),
### returns a shifted reference label
def getShiftedReference(originalLabel, vector):
    if type(vector) == tuple:
        vector = list(vector)
    elif type(vector) == list:
        vector = copy.copy(vector)
    colAbsolute = "" # empty strings
    rowAbsolute = "" # will be "$" if an absolute ref is needed
    if originalLabel[0] == "$":
        colAbsolute = "$"
        vector[1] = 0 # no col shift, col ref is absolute
    if "$" in originalLabel[1:]:
        rowAbsolute = "$"
        vector[0] = 0 # no row shift, row ref is absolute
    usableLabel = originalLabel.replace("$", "")
    usableNumber = getCellNumber(usableLabel)
    return (colAbsolute + getColumnLabel(usableNumber[1] + vector[1]) +
            rowAbsolute + str(usableNumber[0] + vector[0]))

## getVector
### get shift vector (list),
### given origin and destination labels (string)
def getVector(origin, destination):
    originLoc = getCellNumber(origin)
    destLoc = getCellNumber(destination)
    return [destLoc[0]-originLoc[0],
            destLoc[1]-originLoc[1]]

## translateFormula
### given an origin label (string),
### a destination label (string),
### and an input string,
### returns a shifted version of the input string
### with adjusted formula references
def translateFormula(origin, destination, inputStr):
    vector = getVector(origin, destination)
    labelPattern = re.compile("\$?[a-zA-Z]+\$?[0-9]+")
    inputList = inputStr.split('"')
    resultStr = ""
    for i in xrange(len(inputList)):
        references = None
        if ((i % 2) != 0): # odd numbered - surrounded by "'s
            resultStr += '"' + inputList[i] + '"'
        else: # fun case - replace references
            workingStr = inputList[i]
            references = labelPattern.findall(workingStr)
            for reference in references:
                if (len(reference) == 0):
                    continue
                else:
                    refStart = workingStr.find(reference)
                    refEnd = refStart + len(reference)
                    resultStr += workingStr[0:refStart]
                    workingStr = workingStr[refEnd:]
                    # last index prior to ref (or 0, if ref starts at 0)
                    updatedRef = getShiftedReference(reference, vector)
                    resultStr += updatedRef
            resultStr += workingStr
    return resultStr