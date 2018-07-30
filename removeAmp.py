from tqdm import tqdm
import openpyxl
import pygame
import re
import sys
import time

printing = True
saving = True

# Strip punctuation and lowercase a string
# standardize_str(word)
#{{{
def standardize_str(word):
    # print(word)
    if word:
        word = str(word).strip()
        word = re.sub('&amp;', '&', word)
        word = re.sub(';$', '', word)
    return word
#}}}za

# number_from_column(column_letter)
# {{{
def number_from_column(column_letter):
    return ord(column_letter) - 64
#}}}

# removeAmps()
#{{{
def removeAmps():
    if printing:
        print("Opening...")

    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    filename = args[0]

    # if len(args) == 0:
        # print(

    wb = openpyxl.load_workbook(filename)
    sheet = wb.worksheets[0]

    first = 2
    last = sheet.max_row + 1
    changes = 0

    for row in tqdm(range(first, last)):
        for cell in sheet[str(row)]:
            word = cell.value
            if word != '' and word != 'None' and word != 'Null':
                formatted = standardize_str(word)
                if word != formatted:
                    changes = changes + 1
                    print(cell.column + str(row), word, "->", formatted)
                if saving:
                    if formatted:
                        sheet[cell.column + str(row)].value = str(formatted).strip()

    if printing:
        print("Processed " + str((last - first) * len(sheet['1'])) + " rows...")
        print("Changed   " + str(changes) + " values...")

    # add the word 'formatted' and save the new file where the original is
    newName = 'amps'
    index = filename[::-1].find('/')
    if printing and saving:
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

removeAmps()

'''
###########
# TESTING #
###########
test1 = 'Andrew;'
result1 = standardize_str(test)
test2 = 'Andrew'
result2 = standardize_str(test2)
test3 = 'Andrew &amp; Evan'
result3 = standardize_str(test3)
print(test, '->', result)
print(test2, '->', result2)
print(test3, '->', result3)
'''
