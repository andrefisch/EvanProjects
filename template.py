import openpyxl
import os
import pygame
import sys
import time

# GETS ALL COLUMNS FROM THIS ROW
# print(sheet[1])

def template():
    # Get all files of a specific type in this directory
    files = [x for x in os.listdir() if x.endswith(".xlsx")]
    for eachfile in files:
        print("----------------------")
        print(eachfile)
        print("----------------------")

    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    fileName = args[0]
    numberArg = int(args[1])
    cols = args[2:]

    # Open a file with sys.argv
    with open(sys.argv[1]) as f:
        print(f)

    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("transactions")

    # if 'sheet' appears randomly we can delete it
    rm = out['Sheet']
    out.remove(rm)

    #################
    # DO STUFF HERE #
    #################
    first = 2
    last = outsheet.max_row + 1
    for col in cols:
        for row in range (start, sheet.max_row + 1):
            outsheet['A1'].value = "DATA GOES HERE"


    print("Processed " + str((last - first) * len(cols)) + " rows...")
    print("Changed   " + str(changes) + " values...")

    # add the word 'formatted' and save the new file where the original is
    newName = 'new'
    index = fileName[::-1].find('/')
    end = fileName[-index - 1:]
    fileName = fileName[:-index - 1] + newName + end[0].capitalize() + end[1:]
    print("Saving " + fileName)
    wb.save(fileName)





    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName)
    sheet = wb[sheetName]
    sheet = wb.worksheets[0]

    #################
    # DO STUFF HERE #
    #################
    first = 2
    last = outsheet.max_row + 1
    for col in cols:
        for row in range (start, sheet.max_row + 1):
            sheet['A1'].value = "DATA GOES HERE"


    print("Processed " + str((last - first) * len(cols)) + " rows...")
    print("Changed   " + str(changes) + " values...")

    # add the word 'formatted' and save the new file where the original is
    newName = 'better'
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

template()
