import geocoder
import openpyxl
import pygame
import sys
import time

COUNTRY_MATCH    = "A"
ADDED            = "B"
COUNTRY_CODE     = "C"
PORT_CODE        = "D"
DESCRIPTION      = "E"
DESCRIPTION_SANS = "F"
WHO_KNOWS        = "G"
LIST_NUMS        = "H"
MORE_CODES       = "I"
NUM_NUMS         = "J"
NOTHER_NOTES     = "K"
LATLONG          = "L"

def getLatLng(address):
    g = geocoder.google(address)
    return g.latlng

def getCoordinates():
    # Open an existing excel file
    # Uses sys.argv to pass in arguments
    args = sys.argv[1:]
    fileName = args[0]
    if len(args) > 1:
        requests = int(args[1])
    else:
        requests = 2500

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
    first = 37481
    last = sheet.max_row
    # last = 300
    success       = 0
    fail          = 0
    already_had   = 0
    cached        = 0
    no_address    = 0
    requests_sent = 0
    threshold     = requests
    addresses     = {}
    coords        = []
    for row in range (first, last + 1):
        if (threshold <= requests_sent):
            print(str(threshold) + " requests were sent")
            break
        coords = []
        # report = str(row) + ": " + str(format((row - first) / (last - first) * 100.00, '.2f')) + "%: " + sheet[COUNTRY_CODE + str(row)].value + " " + sheet[PORT_CODE + str(row)].value
        report = str(row) + ": " + str(requests_sent + 1) + ": " + str(sheet[COUNTRY_CODE + str(row)].value) + " " + str(sheet[PORT_CODE + str(row)].value)
        address      = sheet[DESCRIPTION + str(row)].value + " port"
        listedCoords = sheet[LATLONG     + str(row)].value
        # If there is an address listed in the spreadsheet
        # check to see if address is in address list
        if address != None and listedCoords == None:
            # if it is, fill in the coordinates
            requests_sent = requests_sent + 1
            report = report + "   ASKING GOOGLE: "
            coords = getLatLng(address)
            # saving every 25 entries gives google a break and prevents loss of data thru crashing
            if requests_sent > 0 and requests_sent % 25 == 0:
                print("Taking a break to save...")
                wb.save("betterFile.xlsx")
            # if we are given valid coordinates
            if coords != None and coords != []:
                success = success + 1
                report = report + "SUCCESS: " 
                report = report + str(coords)
                # write them in the spreadsheet
                sheet[LATLONG  + str(row)].value = str(coords[0]) + " " + str(coords[1])
            # otherwise nothing to do
            else:
                fail = fail + 1
                report = report + "failed..."
        # otherwise nothing to do
        else:
            no_address = no_address + 1

        print(report)

    wb.save("betterFile.xlsx")

    print()
    print("Made               " + str(requests_sent)        + " requests")
    print("Acquired           " + str(success)              + " coordinates")
    print("Failed to acquire  " + str(fail)                 + " coordinates")
    total = requests_sent + success + fail + already_had + no_address
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
