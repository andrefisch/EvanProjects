from datetime import datetime
import openpyxl
import os
import pygame
import re
import sys
import time

# given an incomplete date string (Jun 2018) convert it to a date number -> (06/01/2018)
def fixDate(date):
    if type(date) is str:
        if date[0] == "'":
            date = date[1:]
        if len(date) == 6:
            date = date[:3] + ' 20' + date[4:]
        return datetime.strptime(date, '%b %Y')
        '''
        switcher = {
           "Jan": date[4:] + '-01-01',
           "Feb": date[4:] + '-02-01',
           "Mar": date[4:] + '-03-01',
           "Apr": date[4:] + '-04-01',
           "May": date[4:] + '-05-01',
           "Jun": date[4:] + '-06-01',
           "Jul": date[4:] + '-07-01',
           "Aug": date[4:] + '-08-01',
           "Sep": date[4:] + '-09-01',
           "Oct": date[4:] + '-10-01',
           "Nov": date[4:] + '-11-01',
           "Dec": date[4:] + '-12-01' 
        }
        return switcher.get(date[:3], "Invalid month")
        '''
    else:
        return date

def fixDates():
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
            sheet[col + str(row)].value = fixDate(sheet[col + str(row)].value)

    # Save the file
    wb.save("fixedDates.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

fixDates()
