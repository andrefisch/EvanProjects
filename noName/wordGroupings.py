import openpyxl
import operator
import os
import pygame
import re
import sys
import time

def joinArray(array):
    return ' '.join(array)

def wordGroupings():
    '''
    # Get all files of a specific type in this directory
    files = [x for x in os.listdir() if x.endswith(".xlsx")]
    for eachfile in files:
        print("----------------------")
        print(eachfile)
        print("----------------------")

    # Open a file with sys.argv
    with open(sys.argv[1]) as f:
        print(f)
    '''

    # Uses sys.argv to pass in arguments
    # FILE, TARGET RATING, NUMBER OF WORDS
    args = sys.argv[1:]
    fileName = args[0]
    rating = int(args[1])
    numWords = int(args[2])


    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    #################
    # DO STUFF HERE #
    #################
    '''
    - put all text from relevant reviews in single variable
    - lowercase everything then use regex to split into words
    - use 3-piece array to store three words at a time
    - check dict for existence of key
      - if it doesent exist set it to 1
      - otherwise increment it
    - sort dict by frequency of occurrence
    '''

    start = 2
    end = sheet.max_row + 1

    # GATES
    SETS      = False
    WORDS     = False
    SENTENCES = True
    # Variables
    table = {}
    singleWords = {}
    text = ""
    words = []
    # wordNum = 3
    regexWords = "[\w'-]+"
    regexSentences = "[A-Z].*?[\.?!]"
    RATING = "D"
    REVIEW = "E"

    # put all text from relevant reviews in single variable
    for row in range(start, end):
        # If the review contains the rating of your choice add the review to text variable
        if sheet[RATING + str(row)].value == rating:
            text = text + " " + sheet[REVIEW + str(row)].value

    if SENTENCES:
        keyWords = ["waiter", "never", "worst", "flavorless", "poor", "terrible", "overcooked", "disgusting", "service"]
        # split by sentences
        matchSentences = re.findall(regexSentences, text)
        for i in range(0, len(matchSentences)):
            for j in range(0, len(keyWords)):
                if keyWords[j] in matchSentences[i]:
                    print(matchSentences[i])
                    break

    # lowercase everything then use regex to split into words
    text = text.lower()
    matchWords = re.findall(regexWords, text)

    stopWords = 'i me my myself we our ours ourselves you your yours yourself yourselves ni he him his himself she her hers herself it its itself they them their theirs themselves what which who whom this that these those am is are was were be been being have has had having do does did doing a an the and but if or because as until while of at by for with about against between into through during before after above below to from up down in out on off over under again further then once here there when where why how all any both each few more most other some such no nor not only own same so than too very can will just should now nthe'

    if WORDS:
        # works correctly but doesn't give meaningful data
        # SINGLE WORDS
        # loop through the words and add them to array
        for i in range(0, len(matchWords)):
            # check dict for existence of key
            if matchWords[i] not in stopWords:
                key = matchWords[i]
            else:
                continue
            if key not in table:
                # if it doesent exist set it to 1
                table[key] = 1
            else:
                # otherwise increment it
                table[key] = table[key] + 1

        # sort the dictionary
        sorted_x = sorted(table.items(), key=operator.itemgetter(1))

        print(sorted_x)

    if SETS:
        # works correctly but doesn't give meaningful data
        # SENTENCES
        # loop through the words and add them to array
        for i in range(0, len(matchWords)):
            if len(words) == numWords:
                # check dict for existence of key
                key = joinArray(words)
                if key not in table:
                    # if it doesent exist set it to 1
                    table[key] = 1
                else:
                    # otherwise increment it
                    table[key] = table[key] + 1
                # remove the first word
                del words[0]
            # add the next one if it isnt a stopwords
            # if matchWords[i] not in stopWords:
            words.append(matchWords[i])

        # sort the dictionary
        sorted_x = sorted(table.items(), key=operator.itemgetter(1))

        print(sorted_x)


    wb.save("betterFile.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

wordGroupings()
