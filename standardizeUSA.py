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
    first = 2
    last = sheet.max_row + 1
    changes = 2
    for col in cols:
        for row in range (first, last):
            country = str(sheet[col+ str(row)].value)
            regexUSA = '(\\bU\.?S\.?A?\.?\\b|United ?States ?(of ?America)?)'
            matchUSA = re.search(regexUSA, country, re.IGNORECASE)
            if matchUSA:
                changes = changes + 1
                print(col + str(row) + ": ", country, '->', "United States")
                sheet[col+ str(row)].value = "United States"
            regexUK = '(\\bU\.?K\.?\\b)'
            matchUK = re.search(regexUK, country, re.IGNORECASE)
            if matchUK:
                changes = changes + 1
                print(col + str(row) + ": ", country, '->', "United Kingdom")
                sheet[col+ str(row)].value = "United Kingdom"


    print("Processed " + str((last - first) * len(cols)) + " rows...")
    print("Changed   " + str(changes) + " values...")

    # add the word 'formatted' and save the new file where the original is
    newName = 'countries'
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

    print()
    print("Done!")
#}}}

standardize_USA(sys.argv)
