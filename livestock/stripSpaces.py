from datetime import datetime
import openpyxl
import os
import pygame
import re
import sys
import time

def reformCol(string):
    arr = string.split('/')
    for i in range(0, len(arr)):
        arr[i] = arr[i].strip()
    return "/".join(arr)

def stripSpaces():
    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    print(args)
    fileName = args[0]
    cols = args[1:]

    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    #################
    # DO STUFF HERE #
    #################
    start = 3
    end = sheet.max_row + 1
    for col in cols:
        for row in range(start, end):
            old = sheet[col + str(row)].value
            new = reformCol(old)
            print(col, row, old, new)
            sheet[col + str(row)].value = new

    # Save the file
    wb.save("fixedPorts.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

stripSpaces()
