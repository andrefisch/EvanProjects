import geocoder
import openpyxl
import pygame
import sys
import time

LAST_NAME  = "A"
FIRST_NAME = "B"
STREETNUM  = "H"
STREETNAME = "J"
CITY       = "L"
STATE      = "M"
ZIP        = "N"
LATITUDE   = "AA"
LONGITUDE  = "AB"

def getAddress(sheet, row):
    return str(sheet[STREETNUM   + str(row)].value) + ", " + \
            str(sheet[STREETNAME + str(row)].value) + ", " + \
            str(sheet[CITY       + str(row)].value) + ", " + \
            str(sheet[STATE      + str(row)].value) + ", " + \
            str(sheet[ZIP        + str(row)].value) + ", "

def getLatLng(address):
    g = geocoder.google(address)
    return g.latlng

def getCoordinates():
    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    fileName = args[0]
    if len(args) > 2:
        requests = int(args[1])
    else:
        requests = 2500

    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    #################
    # DO STUFF HERE #
    #################
    '''
    # check to see if there is an address in the spreadsheet
      # check to see if address is in address list
        # if it is check if we have no coordinates
          # fill in the coordinates
        # otherwise if we have coordinates already
          # add address to dictionary
        # otherwise get coordinates from google
          # if we are given valid coordinates
            # write them in the spreadsheet
            # save them in the dictionary
          # otherwise do nothing
      # otherwise nothing to do
    '''
    first = 2
    last = sheet.max_row + 1
    step = 50
    # last = 300
    success       = 0
    fail          = 0
    already_had   = 0
    cached        = 0
    no_address    = 0
    requests_sent = 0
    addresses     = {}
    coords        = []
    for row in range (first, last):
        coords = []
        report = str(row) + ": " + str(requests_sent + 1) + ": " + sheet[FIRST_NAME + str(row)].value + " " + sheet[LAST_NAME + str(row)].value +  "    "
        address     = sheet[STREETNAME  + str(row)].value
        listedCoords = sheet[LATITUDE + str(row)].value
        # If there is an address listed in the spreadsheet
        # check to see if address is in address list
        if address != None:
            # get that address
            address = getAddress(sheet, row)
            # if it is, fill in the coordinates
            if address in addresses:
                cached = cached + 1
                if listedCoords == None:
                    report = report + "   FILLING IN ADDRESS FROM CACHE"
                    coords = addresses[address]
                    sheet[LATITUDE  + str(row)].value = coords[0]
                    sheet[LONGITUDE + str(row)].value = coords[1]
            # otherwise if we have coordinates already
            elif listedCoords != None:
                already_had = already_had + 1
                report = report + "   caching address"
                # add address to dictionary
                addresses[address] = [sheet[LATITUDE + str(row)].value, sheet[LONGITUDE + str(row)].value]
            # otherwise get coordinates from google
            else:
                requests_sent = requests_sent + 1
                report = report + "   ASKING GOOGLE: "
                coords = getLatLng(address)
                # take a quick nap every step queries sent to prevent overload
                if requests_sent > 0 and requests_sent % step == 0:
                    print("Taking a break to save...")
                    wb.save("latLongFile.xlsx")
                    time.sleep(5)
                # if we are given valid coordinates
                if coords != None and coords != []:
                    success = success + 1
                    report = report + "SUCCESS: " 
                    report = report + str(coords)
                    # write them in the spreadsheet
                    sheet[LATITUDE  + str(row)].value = coords[0]
                    sheet[LONGITUDE + str(row)].value = coords[1]
                    # save them in the dictionary
                    addresses[address] = coords
                # otherwise nothing to do
                else:
                    fail = fail + 1
                    report = report + "failed..."
        # otherwise nothing to do
        else:
            no_address = no_address + 1

        print(report)

    wb.save("latLongFile.xlsx")

    print()
    print("Made               " + str(requests_sent)        + " requests")
    print("Acquired           " + str(success)              + " coordinates")
    print("Failed to acquire  " + str(fail)                 + " coordinates")
    print("Already had        " + str(cached + already_had) + " coordinates")
    print("Had no address for " + str(no_address)           + " coordinates")
    total = requests_sent + success + fail + cached + already_had + no_address
    print("TOTAL:             " + str(total))

    if requests_sent > 0:
        print(str(format((success / requests_sent) * 100.00, '.2f')) + "% success rate")
        print(str(format((fail / requests_sent) * 100.00, '.2f')) + "% failure rate")


    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

getCoordinates()
