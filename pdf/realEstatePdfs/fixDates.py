import openpyxl
import os
import pygame
import sys
import time

# GETS ALL COLUMNS FROM THIS ROW
# print(sheet[1])

def fix_date(date):
    month, day, year = date.strip().replace("  ", " ").split(" ")
    months = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}
    if len(day) == 1:
        day = '0' + str(day)
    if ':' in year:
        year = '2018'
    return str(year) + months[month] + str(day)


def fix_dates():
    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    fileName = args[0]
    cols = args[1:]
    # OPEN AN EXISTING EXCEL FILE
    wb = openpyxl.load_workbook(fileName)
    # Usually you just want to use this option
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
            new = fix_date(original)
            if new != original:
                changes = changes + 1
                sheet[col + str(row)].value = new
                print(original, '->', new)


    print("Processed " + str((last - first) * len(cols)) + " rows...")
    print("Changed   " + str(changes) + " values...")

    # add the word 'improved' and save the new file where the original is
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
