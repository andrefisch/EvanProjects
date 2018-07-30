import openpyxl
import os
import pygame
import re
import sys
import time

def combine_notes(roles, jobD, companyD, started, note):
    notes = ""

    if roles:
        notes = "Roles: " + roles + '; '
    if jobD:
        notes = notes + 'Job Description: ' + jobD + '; '
    if companyD:
        notes = notes + 'Company Description: ' + companyD + '; '
    if started:
        notes = notes + 'Started Career in: ' + started + '; '
    if note:
        notes = notes + 'Original Notes: ' + note + '; '

    return notes


def fix_dates():
    # Uses sys.argv to pass in arguments
    args        = sys.argv[1:]
    fileName    = args[0]
    roleCol     = args[1]
    jobDCol     = args[2]
    companyDCol = args[3]
    startedCol  = args[4]
    noteCol     = args[5]

    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    #################
    # DO STUFF HERE #
    #################
    first = 2
    last = sheet.max_row + 1
    for row in range (first, last):
        roles    = sheet[roleCol     + str(row)].value
        jobD     = sheet[jobDCol     + str(row)].value
        companyD = sheet[companyDCol + str(row)].value
        started  = sheet[startedCol  + str(row)].value
        note     = sheet[noteCol     + str(row)].value

        sheet[noteCol + str(row)].value = combine_notes(roles, jobD, companyD, started, note)


    # add the word 'formatted' and save the new file where the original is
    newName = 'notes'
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

fix_dates()
