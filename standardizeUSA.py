import openpyxl
import pygame
import re
import sys
import time

# Change 'US' and 'United States of America' to 'United States'
# standardize_USA(fileName, start, *cols)
#{{{
# def standardize_USA(fileName, column):
def standardize_USA(*args):
    # turn the arguments into variable names
    args = args[0]
    fileName = args[1]
    cols = args[2:]
    # Open an existing excel file
    print("Opening...")
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    #################
    # DO STUFF HERE #
    #################
    start = 2
    for col in cols:
        for row in range (start, sheet.max_row + 1):
            country = str(sheet[col+ str(row)].value)
            regexUSA = '(U\.?S\.?A?\.?|United ?States ?(of ?America)?)'
            matchUSA = re.search(regexUSA, country, re.IGNORECASE)
            if matchUSA:
                sheet[col+ str(row)].value = "United States"
            regexUK = '(U\.?K\.?)'
            matchUK = re.search(regexUK, country, re.IGNORECASE)
            if matchUK:
                sheet[col+ str(row)].value = "United Kingdom"


    print("Saving...")

    wb.save("betterFile.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

    print()
    print("Done!")
#}}}

standardize_USA(sys.argv)
