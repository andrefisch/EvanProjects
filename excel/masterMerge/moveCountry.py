import openpyxl
import os
import pygame
import sys
import time

# GETS ALL COLUMNS FROM THIS ROW
# print(sheet[1])


def moveCountry():
    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    fileName = args[0]

    EXTENDED = args[1]
    COUNTRY  = args[2]

    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    #################
    # DO STUFF HERE #
    #################
    first = 2
    last = sheet.max_row + 1
    changes = 0
    for row in range (first, last):
        if sheet[EXTENDED + str(row)].value and not sheet[COUNTRY + str(row)].value:
            sheet[COUNTRY + str(row)].value = sheet[EXTENDED + str(row)].value
            sheet[EXTENDED + str(row)].value = ""
            changes = changes + 1
            print(str(row) + ": moved", sheet[COUNTRY + str(row)].value)

    print("Processed " + str(last - first) + " rows...")
    print("Changed   " + str(changes) + " values...")

    # add the word 'formatted' and save the new file where the original is
    newName = 'moved'
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

moveCountry()
