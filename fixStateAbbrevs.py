from importDict import importDict as impDi
import openpyxl
import os
import pygame
import sys
import time

# GETS ALL COLUMNS FROM THIS ROW
# print(sheet[1])

def fixStateAbbreviations():
    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    fileName = args[0]
    cols = args[1:]

    dicty = impDi('../../codeToState.txt')

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
            if sheet[col + str(row)].value:
                info = sheet[col + str(row)].value.strip().upper()
                if len(info) == 2:
                    if info in dicty:
                        changes = changes + 1
                        state = dicty[info]
                        sheet[col + str(row)].value = state
                        print(col + str(row) + ":", info, '->', state)

    print("Processed " + str((last - first) * len(cols)) + " rows...")
    print("Changed   " + str(changes) + " values...")

    # add the word 'formatted' and save the new file where the original is
    newName = 'states'
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

fixStateAbbreviations()
