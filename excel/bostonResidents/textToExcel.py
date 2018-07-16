import csv
import openpyxl
import os
import pygame
import re
import sys
import time

# GETS ALL COLUMNS FROM THIS ROW
# print(sheet[1])

def sqlToExcel():
    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("transactions")

    # if 'sheet' appears randomly we can delete it
    rm = out['Sheet']
    out.remove(rm)

    # Open a file with sys.argv
    with open(sys.argv[1]) as f:
        total = 0
        row = 0
        lines = []
        for line in f:
            total = total + 1
            lines.append(line)

        print('total lines:', total)

        for line in range(4, len(lines) - 3):
            row = row + 1
            values = lines[line].strip('\n').split('\t')
            for i in range(0, len(values)):
                outsheet[chr(i + 65) + str(row)].value = values[i]

        # Save the file
        out.save("queryResults.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

sqlToExcel()

