import openpyxl
import os
import pygame
import re
import sys
import time

saving = True
printing = True
testing = True
low_threshold = 75
high_threshold = 100
min_word_len = 5
bad_words = ['korea', 'global', 'shipping', 'golden', 'petra', 'shanghai', 'qingdao', 'marita', 'univer', 'royal', 'prima', 'universal']

# Edit Distance Function
#{{{
# modified edit distance algorithm:
# - see if string one is a substring of the next string
#   - if it is there is a match
#   - it not look at edit distance for substring and substring of equal length
#     - if it is above a certain percentage or below a certain edit number we can safely assume they are a match. keep that string and try with next one

def edit_distance(word1, word2):
    if printing:
        print("Looking at '" + word1 + "' and '" + word2 + "'")
    word1 = word1.lower()
    word2 = word2.lower()
    len_1 = len(word1)
    len_2 = len(word2)

    # make sure shorter word is first
    if len_1 > len_2:
        word1, word2 = word2, word1
        len_1 = len_2

    if (len_1 >= min_word_len and word1 not in bad_words):
        if (word1 in word2):
            if printing:
                print ("Edit distance: 0")
                print ("Percent Match: 100")
            return (0, 100)
        else:
            # shorten longer word to length of first word
            word2 = word2[0:len_1]
            # the matrix whose last element -> edit distance
            x = [[0] * (len_1 + 1) for _ in range(len_1 + 1)]

            # initialization of base case values
            for i in range(0, len_1 + 1): 
                x[i][0] = i
            for j in range(0, len_1 + 1):
                x[0][j] = j
            for i in range (1, len_1 + 1):
                for j in range(1, len_1 + 1):
                    if word1[i - 1] == word2[j - 1]:
                        x[i][j] = x[i - 1][j - 1] 
                    else:
                        x[i][j]= min(x[i][j - 1], x[i - 1][j], x[i - 1][j - 1]) + 1
            edit_distance = x[i][j]
            percent_match = ((len_1 - edit_distance) / len_1) * 100
            if printing:
                print ("Edit distance " + str(x[i][j]))
                print ("Percent match: " + "%.2f" % percent_match)
            return (edit_distance, percent_match)
    return (0, 0)


#}}}

def fix_street_name(string):
    names = ['street\\b', 'avenue\\b', 'boulevard\\b', 'road\\b', 'place\\b', 'square\\b', 'highway\\b', 'broadway\\b']
    string = string.replace('-st-', '-street-')
    string = string.replace('-rd-', '-road-')
    string = string.replace('-ave-', '-avenue-')
    for name in names:
        regex = '.*' + name
        regex = '(.*' + name + ')'
        match = re.search(regex, string, re.IGNORECASE)
        if match:
            return match.group(1)
    return string


def extrapolate_parent_companies():
    # Uses sys.argv to pass in arguments
    args     = sys.argv[1:]
    fileName = args[0]
    fromCol  = args[1]
    toCol    = args[2]

    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    first = 2
    last = sheet.max_row + 1
    # first = 70
    # last = 1005

    DAUGHTER = fromCol
    PARENT   = toCol

    # all rows algorithm
    # - the first row is automatically its own company
    # - for all other rows take current company we are looking at and compare edit distance to it
    #   - if distance is way off store new company and keep going

    curr = ""
    count = 0
    for row in range (first, last):
        # for first row we take the name automatically
        if row == first:
            curr = sheet[DAUGHTER + str(row)].value
            if testing:
                print("now in first row")
                print(sheet[DAUGHTER + str(row)].value)
            # write the data
            if saving:
                sheet[PARENT + str(row)] = fix_street_name(curr).replace('-', ' ').replace('.pdf', '').title()
        else:
            print(row)
            edit, match = edit_distance(curr, sheet[DAUGHTER + str(row)].value)
            if match > low_threshold and match <= high_threshold:
                sheet[PARENT + str(row)] = fix_street_name(curr).replace('-', ' ').replace('.pdf', '').title()
                count = count + 1
            else:
                curr = sheet[DAUGHTER + str(row)].value
                sheet[PARENT + str(row)] = fix_street_name(curr).replace('-', ' ').replace('.pdf', '').title()

    if printing:
        print()
        print("Out of " + str(last - first) + " companies there were " + str(count) + " duplicates")
        print("Saving...")

    if printing:
        print()
        print("Done!")

    # add the word 'formatted' and save the new file where the original is
    newName = 'better'
    index = fileName[::-1].find('/')
    end = fileName[-index - 1:]
    fileName = fileName[:-index - 1] + newName + end[0].capitalize() + end[1:]
    print("Saving " + fileName)
    wb.save(fileName)

    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

extrapolate_parent_companies()
