from lxml import html
import openpyxl
import pygame
import re
import requests
import string
import sys
import time


def scrapePropertyDetails():
    query = sys.argv[1]
    query = query.replace(' ', '+')
    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("reviews")

    # if 'sheet' appears randomly we can delete it
    rm = out['Sheet']
    out.remove(rm)

    # Yelp unique url endings for each property
    start_urls = ['https://www.cityofboston.gov/assessing/search/' + '?q=' + str(query)]
    num_reviews = 1500 # Number of reviews you want
    page_order = range(0, (num_reviews + 1), 20)
    count = 2
    pageNo = 0

    PARCEL_ID                  = "A"
    ADDRESS                    = "B"
    PROPERTY_TYPE              = "C"
    CLASSIFICATION_CODE        = "D"
    LOT_SIZE                   = "E"
    LIVING_AREA                = "F"
    OWNER2018                  = "G"
    OWNER_ADDRESS              = "H"
    RESIDENTIAL_EXEMPTION      = "I"
    PERSONAL_EXEMPTION         = "J"
    BUILDING_VALUE_2018        = "K"
    LAND_VALUE_2018            = "L"
    TOTAL_VALUE_2018           = "M"
    SQUARE_FEET_OF_LIVING_AREA = "N"
    BASE_FLOOR                 = "O"
    FIREPLACES                 = "P"
    NUMBER_OF_FLOORS           = "Q"
    TOTAL_ROOMS                = "R"
    BEDROOMS                   = "S"
    BATHROOMS                  = "T"
    HALF_BATHROOMS             = "U"
    BATH_STYLE_1               = "V"
    BATH_STYLE_2               = "W"
    BATH_STYLE_3               = "X"
    KITCHEN_STYLE              = "Y"
    KITCHEN_TYPE               = "Z"
    HEAT_TYPE                  = "AA"
    INTERIOR_CONDITION         = "AB"
    INTERIOR_FINISH            = "AC"
    ORIENTATION                = "AD"
    CORNER_UNIT                = "AE"
    VIEW                       = "AF"

    outsheet[PARCEL_ID + '1'].value                  = "PARCEL_ID"
    outsheet[ADDRESS + '1'].value                    = "ADDRESS"
    outsheet[PROPERTY_TYPE + '1'].value              = "PROPERTY_TYPE"
    outsheet[CLASSIFICATION_CODE + '1'].value        = "CLASSIFICATION_CODE"
    outsheet[LOT_SIZE + '1'].value                   = "LOT_SIZE"
    outsheet[LIVING_AREA + '1'].value                = "LIVING_AREA"
    outsheet[OWNER2018 + '1'].value                  = "OWNER2018"
    outsheet[OWNER_ADDRESS + '1'].value              = "OWNER_ADDRESS"
    outsheet[RESIDENTIAL_EXEMPTION + '1'].value      = "RESIDENTIAL_EXEMPTION"
    outsheet[PERSONAL_EXEMPTION + '1'].value         = "PERSONAL_EXEMPTION"
    outsheet[BUILDING_VALUE_2018 + '1'].value        = "BUILDING_VALUE_2018"
    outsheet[LAND_VALUE_2018 + '1'].value            = "LAND_VALUE_2018"
    outsheet[TOTAL_VALUE_2018 + '1'].value           = "TOTAL_VALUE_2018"
    outsheet[SQUARE_FEET_OF_LIVING_AREA + '1'].value = "SQUARE_FEET_OF_LIVING_AREA"
    outsheet[BASE_FLOOR + '1'].value                 = "BASE_FLOOR"
    outsheet[FIREPLACES + '1'].value                 = "FIREPLACES"
    outsheet[NUMBER_OF_FLOORS + '1'].value           = "NUMBER_OF_FLOORS"
    outsheet[TOTAL_ROOMS + '1'].value                = "TOTAL_ROOMS"
    outsheet[BEDROOMS + '1'].value                   = "BEDROOMS"
    outsheet[BATHROOMS + '1'].value                  = "BATHROOMS"
    outsheet[HALF_BATHROOMS + '1'].value             = "HALF_BATHROOMS"
    outsheet[BATH_STYLE_1 + '1'].value               = "BATH_STYLE_1"
    outsheet[BATH_STYLE_2 + '1'].value               = "BATH_STYLE_2"
    outsheet[BATH_STYLE_3 + '1'].value               = "BATH_STYLE_3"
    outsheet[KITCHEN_STYLE + '1'].value              = "KITCHEN_STYL"
    outsheet[KITCHEN_TYPE + '1'].value               = "KITCHEN_TYPE"
    outsheet[HEAT_TYPE + '1'].value                  = "HEAT_TYPE"
    outsheet[INTERIOR_CONDITION + '1'].value         = "INTERIOR_CONDITION"
    outsheet[INTERIOR_FINISH + '1'].value            = "INTERIOR_FINISH"
    outsheet[ORIENTATION + '1'].value                = "ORIENTATION"
    outsheet[CORNER_UNIT + '1'].value                = "CORNER_UNIT"
    outsheet[VIEW + '1'].value                       = "VIEWF"

    for ur in start_urls:
        for o in page_order:
            pageNo = pageNo + 1
            print("Currently on page number " + str(pageNo))
            page = requests.get(ur + ("?start=%s" % o))
            tree = html.fromstring(page.text)
            text = html.tostring(tree).decode("utf-8")
            print(text)
            links = []
            regexLink = '\?=pid[0-9]+'
            for i in range (0, len(text)):
                matchLink = re.search(regexLink, text[i])
                if matchLink:
                    links.matchLinks.group(0)

            print(links)
            break
            '''
            for i in range (0, len(matchReview)):
                outsheet[PARCEL_ID + str(i)].value                  = "PARCEL_ID"
                outsheet[ADDRESS + str(i)].value                    = "ADDRESS"
                outsheet[PROPERTY_TYPE + str(i)].value              = "PROPERTY_TYPE"
                outsheet[CLASSIFICATION_CODE + str(i)].value        = "CLASSIFICATION_CODE"
                outsheet[LOT_SIZE + str(i)].value                   = "LOT_SIZE"
                outsheet[LIVING_AREA + str(i)].value                = "LIVING_AREA"
                outsheet[OWNER2018 + str(i)].value                  = "OWNER2018"
                outsheet[OWNER_ADDRESS + str(i)].value              = "OWNER_ADDRESS"
                outsheet[RESIDENTIAL_EXEMPTION + str(i)].value      = "RESIDENTIAL_EXEMPTION"
                outsheet[PERSONAL_EXEMPTION + str(i)].value         = "PERSONAL_EXEMPTION"
                outsheet[BUILDING_VALUE_2018 + str(i)].value        = "BUILDING_VALUE_2018"
                outsheet[LAND_VALUE_2018 + str(i)].value            = "LAND_VALUE_2018"
                outsheet[TOTAL_VALUE_2018 + str(i)].value           = "TOTAL_VALUE_2018"
                outsheet[SQUARE_FEET_OF_LIVING_AREA + str(i)].value = "SQUARE_FEET_OF_LIVING_AREA"
                outsheet[BASE_FLOOR + str(i)].value                 = "BASE_FLOOR"
                outsheet[FIREPLACES + str(i)].value                 = "FIREPLACES"
                outsheet[NUMBER_OF_FLOORS + str(i)].value           = "NUMBER_OF_FLOORS"
                outsheet[TOTAL_ROOMS + str(i)].value                = "TOTAL_ROOMS"
                outsheet[BEDROOMS + str(i)].value                   = "BEDROOMS"
                outsheet[BATHROOMS + str(i)].value                  = "BATHROOMS"
                outsheet[HALF_BATHROOMS + str(i)].value             = "HALF_BATHROOMS"
                outsheet[BATH_STYLE_1 + str(i)].value               = "BATH_STYLE_1"
                outsheet[BATH_STYLE_2 + str(i)].value               = "BATH_STYLE_2"
                outsheet[BATH_STYLE_3 + str(i)].value               = "BATH_STYLE_3"
                outsheet[KITCHEN_STYLE + str(i)].value              = "KITCHEN_STYL"
                outsheet[KITCHEN_TYPE + str(i)].value               = "KITCHEN_TYPE"
                outsheet[HEAT_TYPE + str(i)].value                  = "HEAT_TYPE"
                outsheet[INTERIOR_CONDITION + str(i)].value         = "INTERIOR_CONDITION"
                outsheet[INTERIOR_FINISH + str(i)].value            = "INTERIOR_FINISH"
                outsheet[ORIENTATION + str(i)].value                = "ORIENTATION"
                outsheet[CORNER_UNIT + str(i)].value                = "CORNER_UNIT"
                outsheet[VIEW + str(i)].value                       = "VIEWF"
                '''


    # Save the file
    out.save(".xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

scrapePropertyDetails()
