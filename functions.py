# functions.py
## classless utility functions

from datetime import datetime

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