import openpyxl
import os
import pygame
import sys
import time

# GETS ALL COLUMNS FROM THIS ROW
# print(sheet[1])

def streetCleaning():
    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    fileName = args[0]

    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    ADDRESS1  = "O"
    ADDRESS11 = "Q"
    ADDRESS12 = "R"
    ADDRESS13 = "S"

    #################
    # DO STUFF HERE #
    #################
    start = 2
    end = sheet.max_row + 1
    for row in range(start, end):
        part1 = sheet[ADDRESS11 + str(row)].value 
        if part1 == None:
            part1 = ''
        sheet[ADDRESS11 + str(row)].value = ''
        part2 = sheet[ADDRESS12 + str(row)].value 
        if part2 == None:
            part2 = ''
        sheet[ADDRESS12 + str(row)].value = ''
        part3 = sheet[ADDRESS13 + str(row)].value 
        if part3 == None:
            part3 = ''
        sheet[ADDRESS13 + str(row)].value = ''

        sheet[ADDRESS1 + str(row)].value = str(part1) + " " + str(part2) + " " + str(part3)

    wb.save("streetCleanedMOTU.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

streetCleaning()
