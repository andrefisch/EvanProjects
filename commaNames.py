import openpyxl
import pygame
import re
import string
import sys
import time

printing = True
saving = True

# Strip punctuation and lowercase a string
# standardize_str(word)
#{{{
def standardize_str(word):
    word = word.title().strip()
    if word and ',' in word:
        index = word.find(',')
        name = word[index + 1:] + " " + word[:index]
        return name.strip()
    else:
        return word
#}}}

# number_from_column(column_letter)
# {{{
def number_from_column(column_letter):
    return ord(column_letter) - 64
#}}}

#{{{
def comma_names():
    if printing:
        print("Opening...")

    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    filename = args[0]
    cols = args[1:]

    # if len(args) == 0:
        # print(

    wb = openpyxl.load_workbook(filename)
    sheet = wb.worksheets[0]

    first = 2
    last = sheet.max_row + 1
    changes = 0

    for col in cols:
        for row in range (first, last):
            word = str(sheet[col + str(row)].value)
            if word != '' and word != 'None' and word != 'Null':
                formatted = standardize_str(word)
                if word != formatted:
                    changes = changes + 1
                    print(col + str(row) + ": ", word, '->', formatted)

                if saving:
                    sheet[col + str(row)].value = formatted.strip()

    if printing:
        print("Processed " + str((last - first) * len(cols)) + " rows...")
        print("Changed   " + str(changes) + " values...")

    if printing and saving:
        # add the word 'formatted' and save the new file where the original is
        newName = 'reversed'
        index = filename[::-1].find('/')
        end = filename[-index - 1:]
        filename = filename[:-index - 1] + newName + end[0].capitalize() + end[1:]
        print("Saving " + filename)
        wb.save(filename)

    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()
#}}}

comma_names()
