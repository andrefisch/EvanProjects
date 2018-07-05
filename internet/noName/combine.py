import openpyxl
import os
import pygame
import sys
import time

def template():
    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("transactions")
    # Open an existing excel file
    wb1 = openpyxl.load_workbook(sys.argv[1])
    sheet1 = wb1.worksheets[0]
    # Open an existing excel file
    wb2 = openpyxl.load_workbook(sys.argv[2])
    sheet2 = wb2.worksheets[0]

    # if 'sheet' appears randomly we can delete it
    rm = out['Sheet']
    out.remove(rm)

    #################
    # DO STUFF HERE #
    #################
    DATE   = "A"
    STARS  = "B"
    REVIEW = "C"
    SOURCE = "D"
    outsheet[DATE   + '1'].value = "DATE"
    outsheet[STARS  + '1'].value = "STARS"
    outsheet[REVIEW + '1'].value = "REVIEW"
    outsheet[SOURCE + '1'].value = "SOURCE"
    
    row = 2
    for i in range(2, sheet1.max_row + 1):
        outsheet[DATE   + str(row)].value = sheet1[DATE   + str(i)].value
        outsheet[STARS  + str(row)].value = sheet1[STARS  + str(i)].value
        outsheet[REVIEW + str(row)].value = sheet1[REVIEW + str(i)].value
        outsheet[SOURCE + str(row)].value = sheet1[SOURCE + str(i)].value
        row = row + 1

    for i in range(2, sheet2.max_row + 1):
        outsheet[DATE   + str(row)].value = sheet2[DATE   + str(i)].value
        outsheet[STARS  + str(row)].value = sheet2[STARS  + str(i)].value
        outsheet[REVIEW + str(row)].value = sheet2[REVIEW + str(i)].value
        outsheet[SOURCE + str(row)].value = sheet2[SOURCE + str(i)].value
        row = row + 1

    out.save("combined.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

template()
