import sys

def number_from_column(column_letter):
    return ord(column_letter) - 64

def importDict(fileName):
    # Open a file with sys.argv
    f = open(fileName, 'r')
    dicty = {}
    for line in f:
        # split line by semicolon
        key, value = line.split(';')
        # do not take the \n at the end
        dicty[key] = value[:-1]
    return dicty
