from datetime import datetime
from PyPDF2 import PdfFileReader, PdfFileWriter
import math
import openpyxl
import os
import pygame, time
import re
import sys
import time

COMPANY_NAME   = "A"
PERSON_NAME    = "B"
TITLE          = "C"
ADDRESS        = "D"
CITY_STATE_ZIP = "E"
DATE           = "F"

saving = True

# Turn a single PDF page into a list, separated by line breaks
# parsePage(pdf, num)
#{{{
def parsePage(pdf, num):
    with open(pdf, 'rb') as f:
        reader = PdfFileReader(f)
        contents = reader.getPage(num).extractText().split('\n')
        return contents
#}}}

def extractNames():
    if saving:
        # Open the file for editing
        wb = openpyxl.Workbook()
        # Open the worksheet we want to edit
        sheet = wb.create_sheet("ndas")

        # if 'sheet' appears randomly we can delete it
        rm = wb.get_sheet_by_name('Sheet')
        wb.remove_sheet(rm)

        sheet[COMPANY_NAME    + '1'] = "Company Name"
        sheet[PERSON_NAME     + '1'] = "Signee Name"
        sheet[TITLE           + '1'] = "Title"
        sheet[ADDRESS         + '1'] = "Address"
        sheet[CITY_STATE_ZIP  + '1'] = "City, State, Zip"
        sheet[DATE            + '1'] = "Date"

    # Get all files of a specific type in this directory
    files = [x for x in os.listdir() if x.endswith(".pdf")]
    count = 2
    for eachfile in files:
        print("----------------------")
        print(eachfile)
        print("----------------------")
        with open(eachfile, 'rb') as input_file:
            reader = PdfFileReader(input_file)
            pages = reader.getNumPages()

        # Turn all pages into an array
        contents = []
        for i in range (0, pages):
            contents = contents + parsePage(eachfile, i)

        info = "".join(contents[-21:])
        if len(info) > 2000:
            info = info[1000:]
        print(len(info))
        regexInfo = "Veson Nautical Corporation(.*?)By:.*Name:.*Name:(.*)Title:.*Title:(.*)Address:.*Address:(.*)City, State, Zip:.*City, State, Zip:(.*)?Date:(.*)Date:.*"
        matchInfo = re.search(regexInfo, info)
        if matchInfo:
            print(matchInfo.group(0))
            print("Company:          " + matchInfo.group(1))
            print("Name:             " + matchInfo.group(2))
            print("Title:            " + matchInfo.group(3))
            print("Address:          " + matchInfo.group(4))
            print("City, State, Zip: " + matchInfo.group(5))
            print("Date              " + matchInfo.group(6))
            if saving:
                sheet[COMPANY_NAME    + str(count)] = matchInfo.group(1).strip()
                sheet[PERSON_NAME     + str(count)] = matchInfo.group(2).strip().title()
                sheet[TITLE           + str(count)] = matchInfo.group(3).strip().title()
                sheet[ADDRESS         + str(count)] = matchInfo.group(4).strip().title()
                sheet[CITY_STATE_ZIP  + str(count)] = matchInfo.group(5).strip().title()
                date = matchInfo.group(6).strip()
                regexDate = '(...).*?(\d\d?).*?(\d{4})'
                matchDate = re.search(regexDate, date)
                if matchDate:
                    date = matchDate.group(1) + " " + matchDate.group(2) + " " + matchDate.group(3)
                    sheet[DATE            + str(count)] = datetime.strptime(date, '%b %d %Y')
                else:
                    sheet[DATE            + str(count)] = date

        count = count + 1

    if saving:
        wb.save("ndas.xlsx")
        # LMK when the script is done
        pygame.init()
        pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
        pygame.mixer.music.play()
        time.sleep(5)
        pygame.mixer.music.stop()

extractNames()
