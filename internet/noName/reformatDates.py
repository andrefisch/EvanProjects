import openpyxl
import os
import pygame
import re
import sys
import time

def fix_date(string):
    months = {'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10', 'November': '11', 'December': '12'}
    regexNum = '^([0-9]+) (.*?) '
    matchNum = re.search(regexNum, string)
    if matchNum:
        dd = matchNum.group(1)
        mm = matchNum.group(2)
        return months[mm] + '/' + dd + '/' + '2018'
    else:
        return string


def fix_dates():
    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    fileName = args[0]
    cols = args[1:]

    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    #################
    # DO STUFF HERE #
    #################
    first = 2
    last = sheet.max_row + 1
    changes = 0
    for col in cols:
        for row in range (first, last):
            original = sheet[col + str(row)].value
            if original:
                new = fix_date(sheet[col + str(row)].value)
                if original != new:
                    changes = changes + 1
                    print(col + str(row), original, '->', new)
                    sheet[col + str(row)].value = new


    print("Processed " + str((last - first) * len(cols)) + " rows...")
    print("Changed   " + str(changes) + " values...")

    # add the word 'formatted' and save the new file where the original is
    newName = 'dates'
    index = fileName[::-1].find('/')
    end = fileName[-index - 1:]
    fileName = fileName[:-index - 1] + newName + end[0].capitalize() + end[1:]
    print("Saving " + fileName)
    wb.save(fileName)

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

fix_dates()
