import openpyxl
import os
import pygame
import sys
import time

def template():
    # Get all files of a specific type in this directory
    files = [x for x in os.listdir() if x.endswith(".xlsx")]
    for eachfile in files:
        print("----------------------")
        print(eachfile)
        print("----------------------")

    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    firstArg = args[0]
    numberArg = int(args[1])
    otherArgs = args[2:]

    # Open a file with sys.argv
    with open(sys.argv[1]) as f:
        print(f)

    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("transactions")

    # if 'sheet' appears randomly we can delete it
    rm = out.get_sheet_by_name('Sheet')
    out.remove_sheet(rm)

    #################
    # DO STUFF HERE #
    #################
    outsheet['A1'].value = "DATA GOES HERE"

    # Save the file
    out.save("newFile.xlsx")






    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName + ".xlsx")
    sheet = wb[sheetName]

    #################
    # DO STUFF HERE #
    #################
    sheet['A1'].value = "DATA GOES HERE"

    wb.save("betterFile.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

def testing():
    return sys.argv

print(testing())
