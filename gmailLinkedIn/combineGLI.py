# which column is the unique identifier (primary key)
# in event of two names having different data in the same column (collision) should I take most recent?
# you have some accounts which only seem to have multiple rows because they are present in more than one industry or segment. do you want to have all of that information represented in the same cell?

# want to eliminate duplicate companies but keep all information
# eliminate duplicate companies and people

# for first part use contact as unique identifier
# for second part use name as unique identifier

from datetime import datetime
from splitNames import determine_names
import openpyxl
import pygame
import re
import string
import sys
import time

saving = True
printing = True
testing = True
low_threshold = 85
high_threshold = 100
min_word_len = 5

FIRST_NAME           = 'A'
MIDDLE_NAME          = 'B'
LAST_NAME            = 'C'
TITLE                = 'D'
SUFFIX               = 'E'
WEB_PAGE             = 'F'
NOTES                = 'G'
EMAIL_ADDRESS        = 'H'
EMAIL_ADDRESS2       = 'I'
EMAIL_ADDRESS3       = 'J'
HOME_PHONE           = 'K'
MOBILE_PHONE         = 'L'
HOME_ADDRESS         = 'M'
HOME_STREET          = 'N'
HOME_CITY            = 'O'
HOME_STATE           = 'P'
HOME_POSTAL_CODE     = 'Q'
HOME_COUNTRY         = 'R'
CONTACT_MAIN_PHONE   = 'S'
BUSINESS_PHONE       = 'T'
BUSINESS_PHONE2      = 'U'
BUSINESS_FAX         = 'V'
COMPANY              = 'W'
JOB_TITLE            = 'X'
DEPARTMENT           = 'Y'
OFFICE_LOCATION      = 'Z'
BUSINESS_ADDRESS     = 'AA'
BUSINESS_STREET      = 'AB'
BUSINESS_CITY        = 'AC'
BUSINESS_STATE       = 'AD'
BUSINESS_POSTAL_CODE = 'AE'
BUSINESS_COUNTRY     = 'AF'
CATEGORIES           = 'AG'
CONNECTED_ON         = 'AH'
LATITUDE             = 'AI'
LONGITUDE            = 'AJ'


##################
# HELPER METHODS #
##################

# Strip punctuation and lowercase a string
# standardize_str(word)
#{{{
punctuationTable = str.maketrans({key: None for key in string.punctuation})

def standardize_str(word):
    if word != None:
        return word.lower().translate(punctuationTable)
    else:
        return ""
#}}}

# Combine with slash
# special_combine(word1, word2)
# {{{
def special_combine(word1, word2):
    if word1 == "" or word1 == None:
        return word2
    elif word2 == "" or word2 == None:
        return word1
    elif word1 in word2:
        return word2
    elif word2 in word1:
        return word1
    else:
        return word1 + "/" + word2
#}}}

# Add to notes
# add_to_notes(category, info)
#{{{
def add_to_notes(contact, category, info, oldInfo):
    if info != "" and info != None and oldInfo != "" and oldInfo != None and info != oldInfo:
        return str(contact.notes) + "; " + str(category) + " used to be " + str(info)
    else:
        return str(contact.notes)
#}}}

# Keep the non-none value
# keep_non_none(var1, var2)
#{{{
def keep_non_none(var1, var2):
    if var1 == None or var1 == "":
        return (var2, var1)
    elif var2 == None or var2 == "":
        return (var1, var2)
    else:
        return (var2, var1)
#}}}

# Edit Distance Function
# edit_distance(word1, word2, low_threshold, high_threshold)
#{{{
# modified edit distance algorithm:
# - see if string one is a substring of the next string
#   - if it is there is a match
#   - it not look at edit distance for substring and substring of equal length
#     - if it is above a certain percentage or below a certain edit number we can safely assume they are a match. keep that string and try with next one

def edit_distance(word1, word2, low_threshold, high_threshold):
    if printing:
        print("Looking at '" + word1 + "' and '" + word2 + "'")
    word1 = word1.lower()
    word2 = word2.lower()
    len_1 = len(word1)
    len_2 = len(word2)
    edit_distance, percent_match = 0, 0

    # make sure shorter word is first
    if len_1 > len_2:
        word1, word2 = word2, word1
        len_1 = len_2

    if (len_1 >= min_word_len):
        if (word1 in word2):
            if printing:
                print ("Edit distance: 0")
                print ("Percent Match: 100")
            edit_distance, percent_match = 0, 100
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
    if percent_match > low_threshold and percent_match <= high_threshold:
        if printing:
            print("MATCH!")
        return True
    else:
        return False
#}}}

# Contact object with all relevant fields
# Contact
#{{{
class Contact(object):
    first_name           = ''
    middle_name          = ''
    last_name            = ''
    title                = ''
    suffix               = ''
    web_page             = ''
    notes                = ''
    email_address        = ''
    email_address2       = ''
    email_address3       = ''
    home_phone           = ''
    mobile_phone         = ''
    home_address         = ''
    home_street          = ''
    home_city            = ''
    home_state           = ''
    home_postal_code     = ''
    home_country         = ''
    contact_main_phone   = ''
    business_phone       = ''
    business_phone2      = ''
    business_fax         = ''
    company              = ''
    job_title            = ''
    department           = ''
    office_location      = ''
    business_address     = ''
    business_street      = ''
    business_city        = ''
    business_state       = ''
    business_postal_code = ''
    business_country     = ''
    categories           = ''
    connected_on         = ''
    latitude             = ''
    longitude            = ''

    def __init__(self):
        last_name = "-"
        notes     = ""

#}}}

# Create a new contact from sheet information
# new_contact_from_sheet(sheet, row)
#{{{
def new_contact_from_sheet(sheet, row):
    contact = Contact()

    if sheet[FIRST_NAME + str(row)].value != None:
        contact.first_name = sheet[FIRST_NAME + str(row)].value
    if sheet[MIDDLE_NAME + str(row)].value != None:
        contact.middle_name = sheet[MIDDLE_NAME + str(row)].value
    if sheet[LAST_NAME + str(row)].value != None:
        contact.last_name = sheet[LAST_NAME + str(row)].value
    if sheet[SUFFIX + str(row)].value != None:
        contact.suffix = sheet[SUFFIX + str(row)].value
    if sheet[WEB_PAGE + str(row)].value != None:
        contact.web_page = sheet[WEB_PAGE + str(row)].value
    if sheet[NOTES + str(row)].value != None:
        contact.notes = sheet[NOTES + str(row)].value
    if sheet[EMAIL_ADDRESS + str(row)].value != None:
        contact.email_address = sheet[EMAIL_ADDRESS + str(row)].value
    if sheet[EMAIL_ADDRESS2 + str(row)].value != None:
        contact.email_address2 = sheet[EMAIL_ADDRESS2 + str(row)].value
    if sheet[EMAIL_ADDRESS3 + str(row)].value != None:
        contact.email_address3 = sheet[EMAIL_ADDRESS3 + str(row)].value
    if sheet[HOME_PHONE + str(row)].value != None:
        contact.home_phone = sheet[HOME_PHONE + str(row)].value
    if sheet[MOBILE_PHONE + str(row)].value != None:
        contact.mobile_phone = sheet[MOBILE_PHONE + str(row)].value
    if sheet[HOME_ADDRESS + str(row)].value != None:
        contact.home_address = sheet[HOME_ADDRESS + str(row)].value
    if sheet[HOME_STREET + str(row)].value != None:
        contact.home_street = sheet[HOME_STREET + str(row)].value
    if sheet[HOME_CITY + str(row)].value != None:
        contact.home_city = sheet[HOME_CITY + str(row)].value
    if sheet[HOME_STATE + str(row)].value != None:
        contact.home_state = sheet[HOME_STATE + str(row)].value
    if sheet[HOME_POSTAL_CODE + str(row)].value != None:
        contact.home_postal_code = sheet[HOME_POSTAL_CODE + str(row)].value
    if sheet[HOME_COUNTRY + str(row)].value != None:
        contact.home_country = sheet[HOME_COUNTRY + str(row)].value
    if sheet[CONTACT_MAIN_PHONE + str(row)].value != None:
        contact.contact_main_phone = sheet[CONTACT_MAIN_PHONE + str(row)].value
    if sheet[BUSINESS_PHONE + str(row)].value != None:
        contact.business_phone = sheet[BUSINESS_PHONE + str(row)].value
    if sheet[BUSINESS_PHONE2 + str(row)].value != None:
        contact.business_phone2 = sheet[BUSINESS_PHONE2 + str(row)].value
    if sheet[BUSINESS_FAX + str(row)].value != None:
        contact.business_fax = sheet[BUSINESS_FAX + str(row)].value
    if sheet[COMPANY + str(row)].value != None:
        contact.company = sheet[COMPANY + str(row)].value
    if sheet[JOB_TITLE + str(row)].value != None:
        contact.job_title = sheet[JOB_TITLE + str(row)].value
    if sheet[DEPARTMENT + str(row)].value != None:
        contact.department = sheet[DEPARTMENT + str(row)].value
    if sheet[OFFICE_LOCATION + str(row)].value != None:
        contact.office_location = sheet[OFFICE_LOCATION + str(row)].value
    if sheet[BUSINESS_ADDRESS + str(row)].value != None:
        contact.business_address = sheet[BUSINESS_ADDRESS + str(row)].value
    if sheet[BUSINESS_STREET + str(row)].value != None:
        contact.business_street = sheet[BUSINESS_STREET + str(row)].value
    if sheet[BUSINESS_CITY + str(row)].value != None:
        contact.business_city = sheet[BUSINESS_CITY + str(row)].value
    if sheet[BUSINESS_STATE + str(row)].value != None:
        contact.business_state = sheet[BUSINESS_STATE + str(row)].value
    if sheet[BUSINESS_POSTAL_CODE + str(row)].value != None:
        contact.business_postal_code = sheet[BUSINESS_POSTAL_CODE + str(row)].value
    if sheet[BUSINESS_COUNTRY + str(row)].value != None:
        contact.business_country = sheet[BUSINESS_COUNTRY + str(row)].value
    if sheet[CONNECTED_ON + str(row)].value != None:
        contact.connected_on = sheet[CONNECTED_ON + str(row)].value
    if sheet[LATITUDE + str(row)].value != None:
        contact.latitude = sheet[LATITUDE + str(row)].value
    if sheet[LONGITUDE + str(row)].value != None:
        contact.longitude = sheet[LONGITUDE + str(row)].value
    
    return contact
#}}}

# Update Contact from sheet information
# update_contact_from_sheet(sheet, row, contact)
#{{{
def update_contact_from_sheet(sheet, row, contact):
    # update these fields by adding a slash
    # INCLUDE CHANGE IN COMPANY AND CHANGE IN POSITION IN THE NOTES
    contact.notes       = special_combine(contact.notes,       sheet[NOTES       + str(row)].value)
    # update all 
    if sheet[FIRST_NAME + str(row)].value != None or contact.first_name != None:
        info = keep_non_none(contact.first_name, sheet[FIRST_NAME + str(row)].value)
        contact.first_name = info[0]
        contact.notes = add_to_notes(contact, "First Name", info[1], info[0])
    if sheet[MIDDLE_NAME + str(row)].value != None or contact.middle_name != None:
        info = keep_non_none(contact.middle_name, sheet[MIDDLE_NAME + str(row)].value)
        contact.middle_name = info[0]
        contact.notes = add_to_notes(contact, "Middle Name", info[1], info[0])
    if sheet[LAST_NAME + str(row)].value != None or contact.last_name != None:
        info = keep_non_none(contact.last_name, sheet[LAST_NAME + str(row)].value)
        contact.last_name = info[0]
        contact.notes = add_to_notes(contact, "Last Name", info[1], info[0])
    if sheet[SUFFIX + str(row)].value != None or contact.suffix != None:
        info = keep_non_none(contact.suffix, sheet[SUFFIX + str(row)].value)
        contact.suffix = info[0]
        contact.notes = add_to_notes(contact, "Suffix", info[1], info[0])
    if sheet[WEB_PAGE + str(row)].value != None or contact.web_page != None:
        info = keep_non_none(contact.web_page, sheet[WEB_PAGE + str(row)].value)
        contact.web_page = info[0]
        contact.notes = add_to_notes(contact, "Web Page", info[1], info[0])
    if sheet[EMAIL_ADDRESS + str(row)].value != None or contact.email_address != None:
        info = keep_non_none(contact.email_address, sheet[EMAIL_ADDRESS + str(row)].value)
        contact.email_address = info[0]
        contact.notes = add_to_notes(contact, "Email Address", info[1], info[0])
    if sheet[EMAIL_ADDRESS2 + str(row)].value != None or contact.email_address2 != None:
        info = keep_non_none(contact.email_address2, sheet[EMAIL_ADDRESS2 + str(row)].value)
        contact.email_address2 = info[0]
        contact.notes = add_to_notes(contact, "Email Address2", info[1], info[0])
    if sheet[EMAIL_ADDRESS3 + str(row)].value != None or contact.email_address3 != None:
        info = keep_non_none(contact.email_address3, sheet[EMAIL_ADDRESS3 + str(row)].value)
        contact.email_address3 = info[0]
        contact.notes = add_to_notes(contact, "Email Address3", info[1], info[0])
    if sheet[HOME_PHONE + str(row)].value != None or contact.home_phone != None:
        info = keep_non_none(contact.home_phone, sheet[HOME_PHONE + str(row)].value)
        contact.home_phone = info[0]
        contact.notes = add_to_notes(contact, "Home Phone", info[1], info[0])
    if sheet[MOBILE_PHONE + str(row)].value != None or contact.mobile_phone != None:
        info = keep_non_none(contact.mobile_phone, sheet[MOBILE_PHONE + str(row)].value)
        contact.mobile_phone = info[0]
        contact.notes = add_to_notes(contact, "Mobile Phone", info[1], info[0])
    if sheet[HOME_ADDRESS + str(row)].value != None or contact.home_address != None:
        info = keep_non_none(contact.home_address, sheet[HOME_ADDRESS + str(row)].value)
        contact.home_address = info[0]
        contact.notes = add_to_notes(contact, "Home Address", info[1], info[0])
    if sheet[HOME_STREET + str(row)].value != None or contact.home_street != None:
        info = keep_non_none(contact.home_street, sheet[HOME_STREET + str(row)].value)
        contact.home_street = info[0]
        contact.notes = add_to_notes(contact, "Home Street", info[1], info[0])
    if sheet[HOME_CITY + str(row)].value != None or contact.home_city != None:
        info = keep_non_none(contact.home_city, sheet[HOME_CITY + str(row)].value)
        contact.home_city = info[0]
        contact.notes = add_to_notes(contact, "Home City", info[1], info[0])
    if sheet[HOME_STATE + str(row)].value != None or contact.home_state != None:
        info = keep_non_none(contact.home_state, sheet[HOME_STATE + str(row)].value)
        contact.home_state = info[0]
        contact.notes = add_to_notes(contact, "Home State", info[1], info[0])
    if sheet[HOME_POSTAL_CODE + str(row)].value != None or contact.home_postal_code != None:
        info = keep_non_none(contact.home_postal_code, sheet[HOME_POSTAL_CODE + str(row)].value)
        contact.home_postal_code = info[0]
        contact.notes = add_to_notes(contact, "Home Postal Code", info[1], info[0])
    if sheet[HOME_COUNTRY + str(row)].value != None or contact.home_country != None:
        info = keep_non_none(contact.home_country, sheet[HOME_COUNTRY + str(row)].value)
        contact.home_country = info[0]
        contact.notes = add_to_notes(contact, "Home Country", info[1], info[0])
    if sheet[CONTACT_MAIN_PHONE + str(row)].value != None or contact.contact_main_phone != None:
        info = keep_non_none(contact.contact_main_phone, sheet[CONTACT_MAIN_PHONE + str(row)].value)
        contact.contact_main_phone = info[0]
        contact.notes = add_to_notes(contact, "Contact Main Phone", info[1], info[0])
    if sheet[BUSINESS_PHONE + str(row)].value != None or contact.business_phone != None:
        info = keep_non_none(contact.business_phone, sheet[BUSINESS_PHONE + str(row)].value)
        contact.business_phone = info[0]
        contact.notes = add_to_notes(contact, "Business Phone", info[1], info[0])
    if sheet[BUSINESS_PHONE2 + str(row)].value != None or contact.business_phone2 != None:
        info = keep_non_none(contact.business_phone2, sheet[BUSINESS_PHONE2 + str(row)].value)
        contact.business_phone2 = info[0]
        contact.notes = add_to_notes(contact, "Business Phone2", info[1], info[0])
    if sheet[BUSINESS_FAX + str(row)].value != None or contact.business_fax != None:
        info = keep_non_none(contact.business_fax, sheet[BUSINESS_FAX + str(row)].value)
        contact.business_fax = info[0]
        contact.notes = add_to_notes(contact, "Business Fax", info[1], info[0])
    if sheet[COMPANY + str(row)].value != None or contact.company != None:
        info = keep_non_none(contact.company, sheet[COMPANY + str(row)].value)
        contact.company = info[0]
        contact.notes = add_to_notes(contact, "Company", info[1], info[0])
    if sheet[JOB_TITLE + str(row)].value != None or contact.job_title != None:
        info = keep_non_none(contact.job_title, sheet[JOB_TITLE + str(row)].value)
        contact.job_title = info[0]
        contact.notes = add_to_notes(contact, "Job Title", info[1], info[0])
    if sheet[DEPARTMENT + str(row)].value != None or contact.department != None:
        info = keep_non_none(contact.department, sheet[DEPARTMENT + str(row)].value)
        contact.department = info[0]
        contact.notes = add_to_notes(contact, "Department", info[1], info[0])
    if sheet[OFFICE_LOCATION + str(row)].value != None or contact.office_location != None:
        info = keep_non_none(contact.office_location, sheet[OFFICE_LOCATION + str(row)].value)
        contact.office_location = info[0]
        contact.notes = add_to_notes(contact, "Office Location", info[1], info[0])
    if sheet[BUSINESS_ADDRESS + str(row)].value != None or contact.business_address != None:
        info = keep_non_none(contact.business_address, sheet[BUSINESS_ADDRESS + str(row)].value)
        contact.business_address = info[0]
        contact.notes = add_to_notes(contact, "Business Address", info[1], info[0])
    if sheet[BUSINESS_STREET + str(row)].value != None or contact.business_street != None:
        info = keep_non_none(contact.business_street, sheet[BUSINESS_STREET + str(row)].value)
        contact.business_street = info[0]
        contact.notes = add_to_notes(contact, "Business Street", info[1], info[0])
    if sheet[BUSINESS_CITY + str(row)].value != None or contact.business_city != None:
        info = keep_non_none(contact.business_city, sheet[BUSINESS_CITY + str(row)].value)
        contact.business_city = info[0]
        contact.notes = add_to_notes(contact, "Business City", info[1], info[0])
    if sheet[BUSINESS_STATE + str(row)].value != None or contact.business_state != None:
        info = keep_non_none(contact.business_state, sheet[BUSINESS_STATE + str(row)].value)
        contact.business_state = info[0]
        contact.notes = add_to_notes(contact, "Business State", info[1], info[0])
    if sheet[BUSINESS_POSTAL_CODE + str(row)].value != None or contact.business_postal_code != None:
        info = keep_non_none(contact.business_postal_code, sheet[BUSINESS_POSTAL_CODE + str(row)].value)
        contact.business_postal_code = info[0]
        contact.notes = add_to_notes(contact, "Business Postal Code", info[1], info[0])
    if sheet[BUSINESS_COUNTRY + str(row)].value != None or contact.business_country != None:
        info = keep_non_none(contact.business_country, sheet[BUSINESS_COUNTRY + str(row)].value)
        contact.business_country = info[0]
        contact.notes = add_to_notes(contact, "Business Country", info[1], info[0])
    if sheet[CATEGORIES + str(row)].value != None or contact.categories != None:
        info = keep_non_none(contact.categories, sheet[CATEGORIES + str(row)].value)
        contact.categories = info[0]
        contact.notes = add_to_notes(contact, "Categories", info[1], info[0])
    if sheet[CONNECTED_ON + str(row)].value != None or contact.connected_on != None:
        info = keep_non_none(contact.connected_on, sheet[CONNECTED_ON + str(row)].value)
        contact.connected_on = info[0]
        contact.notes = add_to_notes(contact, "Connected On", info[1], info[0])
    if sheet[LATITUDE + str(row)].value != None or contact.latitude != None:
        info = keep_non_none(contact.latitude, sheet[LATITUDE + str(row)].value)
        contact.latitude = info[0]
        contact.notes = add_to_notes(contact, "Latitude", info[1], info[0])
    if sheet[LONGITUDE + str(row)].value != None or contact.longitude != None:
        info = keep_non_none(contact.longitude, sheet[LONGITUDE + str(row)].value)
        contact.longitude = info[0]
        contact.notes = add_to_notes(contact, "Longitude", info[1], info[0])
    return contact
#}}}

################
# MAIN METHODS #
################
# Change 'US' and 'United States of America' to 'United States'
# standardize_USA(fileName, start, column)
#{{{
# def standardize_USA(fileName, start, column):
def standardize_USA(*args):
    # turn the arguments into variable names
    args = args[0]
    fileName = args[1]
    start = int(args[2])
    cols = args[3:]
    # Open an existing excel file
    if printing:
        print("Opening...")
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    #################
    # DO STUFF HERE #
    #################
    for col in cols:
        for row in range (start, sheet.max_row + 1):
            country = str(sheet[col+ str(row)].value)
            regexUSA = '(U\.?S\.?A?\.?|United ?States ?(of ?America)?)'
            matchUSA = re.search(regexUSA, country, re.IGNORECASE)
            if matchUSA:
                sheet[col+ str(row)].value = "United States"
            regexUK = '(U\.?K\.?)'
            matchUK = re.search(regexUK, country, re.IGNORECASE)
            if matchUK:
                sheet[col+ str(row)].value = "United Kingdom"


    if printing:
        print("Saving...")

    wb.save("betterFile.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(3)
    pygame.mixer.music.stop()

    if printing:
        print()
        print("Done!")
#}}}

# Combine LinkedIn and Google Contacts
# Copies all contacts from first list to second
# combine_contacts(fileName1, sheetName1, fileName2, sheetName2)
#{{{
def combine_contacts(fileName1, sheetName1, fileName2, sheetName2):
    if printing:
        print("Opening...")
    wb1 = openpyxl.load_workbook(fileName1 + ".xlsx")
    sheet1 = wb1[sheetName1]
    wb2 = openpyxl.load_workbook(fileName2 + ".xlsx")
    sheet2 = wb2[sheetName2]

    # if 'sheet' appears randomly we can delete it
    # rm = out.get_sheet_by_name('Sheet')
    # out.remove_sheet(rm)

    sheet1[CONNECTED_ON           + '1'].value = "Connected On"

    first = 2
    last = sheet2.max_row
    start = sheet1.max_row

    for row in range (first, start + 1):
        sheet1[CATEGORIES + str(row)].value = str(sheet1[CATEGORIES + str(row)].value) + ":" + str(sheet1['AH' + str(row)].value) + ":" + str(sheet1['AI' + str(row)].value) + ":" + str(sheet1['AJ' + str(row)].value)
        sheet1["AH" + str(row)].value = ""
        sheet1["AI" + str(row)].value = ""
        sheet1["AJ" + str(row)].value = ""

    for row in range (first, last + 1):
        sheet1[FIRST_NAME             + str(row + start)].value = sheet2["A" + str(row)].value
        sheet1[LAST_NAME              + str(row + start)].value = sheet2["B" + str(row)].value
        sheet1[EMAIL_ADDRESS          + str(row + start)].value = sheet2["C" + str(row)].value
        sheet1[COMPANY                + str(row + start)].value = sheet2["D" + str(row)].value
        sheet1[JOB_TITLE              + str(row + start)].value = sheet2["E" + str(row)].value
        sheet1[CATEGORIES             + str(row + start)].value = "LinkedIn Merge 5/14/18"
        sheet1[CONNECTED_ON           + str(row + start)].value = sheet2["F" + str(row)].value

    wb1.save("combined.xlsx")

    if printing:
        print()
        print("Done!")

    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(3)
    pygame.mixer.music.stop()
#}}}

# Combine combined list with a CMAShipping List
# Copies all contacts from sheet 1 to sheet 2
# combine_with_CMA(fileName1, sheetName1, fileName2, sheetName2)
#{{{
'''
- loop through list 1 putting in list 2
  - FORMULA: row + (start + 1) - first
'''
def combine_with_CMA(fileName1, sheetName1, fileName2, sheetName2):
    if printing:
        print("Opening...")
    wb1 = openpyxl.load_workbook(fileName1 + ".xlsx")
    sheet = wb1[sheetName1]
    wb2 = openpyxl.load_workbook(fileName2 + ".xlsx")
    outsheet = wb2[sheetName2]

    first = 3
    last = sheet.max_row
    start = outsheet.max_row

    for row in range (first, last + 1):
        index = row + (start + 1) - first
        if printing:
            print(str(row) + " " + sheet['A' + str(row)].value)
        names = determine_names(sheet['A' + str(row)].value)

        # split name here
        outsheet[FIRST_NAME           + str(index)].value = names['first_name']
        outsheet[MIDDLE_NAME          + str(index)].value = names['middle_name']
        outsheet[LAST_NAME            + str(index)].value = names['last_name']

        # position and company
        outsheet[COMPANY              + str(index)].value = sheet['C' + str(row)].value
        outsheet[JOB_TITLE            + str(index)].value = sheet['B' + str(row)].value

        # put address stuff here
        outsheet[BUSINESS_STREET      + str(index)].value = sheet['D' + str(row)].value
        outsheet[BUSINESS_CITY        + str(index)].value = sheet['G' + str(row)].value
        outsheet[BUSINESS_STATE       + str(index)].value = sheet['H' + str(row)].value
        outsheet[BUSINESS_POSTAL_CODE + str(index)].value = sheet['I' + str(row)].value
        outsheet[BUSINESS_COUNTRY     + str(index)].value = sheet['J' + str(row)].value
        
        # contact info
        outsheet[BUSINESS_PHONE       + str(index)].value = sheet['K' + str(row)].value
        outsheet[EMAIL_ADDRESS        + str(index)].value = sheet['L' + str(row)].value

        # location info
        outsheet[LATITUDE             + str(index)].value = sheet['M' + str(row)].value
        outsheet[LONGITUDE            + str(index)].value = sheet['N' + str(row)].value

        # put CMA SHIPPING
        outsheet[CATEGORIES           + str(index)].value = 'CMA Shipping conference'
        # put March 22, 2018
        outsheet[CONNECTED_ON         + str(index)].value = '03/22/2018'


    wb2.save("combined.xlsx")

    if printing:
        print()
        print("Done!")

    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(3)
    pygame.mixer.music.stop()
#}}}

# Remove Duplicate Contacts
# remove_duplicate_contacts(fileName, sheetName, first, last)
#{{{
# all rows algorithm
# - the first row is automatically its own entity
# - for all other rows take combination of primary contact and current contact we are looking at and compare edit distance to it
#   - if it is a match combine and keep looking through list until we find a bad match
#   - if distance is way off save old contact info. store new contact and keep going
# def remove_duplicate_contacts(fileName, first, match_threshold, *cols):
def remove_duplicate_contacts(*args):
    # turn the arguments into variable names
    args = args[0]
    fileName = args[1]
    first = int(args[2])
    low_threshold = int(args[3])
    cols = args[4:]
    if printing:
        print("Opening...")
    # Open the file for editing
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("contacts")
    wb = openpyxl.load_workbook(fileName)
    sheet = wb.worksheets[0]

    # if 'sheet' appears randomly we can delete it
    rm = out.get_sheet_by_name('Sheet')
    out.remove_sheet(rm)

    # Create a new file to store duplicate contacts
    dupe = openpyxl.Workbook()
    # Open the worksheet we want to edit
    dupesheet = dupe.create_sheet("contacts")
    # if 'sheet' appears randomly we can delete it
    rm = dupe.get_sheet_by_name('Sheet')
    dupe.remove_sheet(rm)

    # - create an object for a new primary contact and account name pair
    #   - store previous object in a new sheet
    #   - store all information here
    #   - look at next contact and see if it matches (edit distance)
    #     - if it is a match combine and keep going, otherwise repeat
    compare = ""
    current = ""
    last = sheet.max_row
    count = 1
    dupes = 2
    contact = Contact()

    # Create Headers
    #{{{
    if saving:
        outsheet[FIRST_NAME           + '1'].value = "First_name"
        outsheet[MIDDLE_NAME          + '1'].value = "Middle_name"
        outsheet[LAST_NAME            + '1'].value = "Last_name"
        outsheet[TITLE                + '1'].value = "Title"
        outsheet[SUFFIX               + '1'].value = "Suffix"
        outsheet[WEB_PAGE             + '1'].value = "Web_page"
        outsheet[NOTES                + '1'].value = "Notes"
        outsheet[EMAIL_ADDRESS        + '1'].value = "Email_address"
        outsheet[EMAIL_ADDRESS2       + '1'].value = "Email_address2"
        outsheet[EMAIL_ADDRESS3       + '1'].value = "Email_address3"
        outsheet[HOME_PHONE           + '1'].value = "Home_phone"
        outsheet[MOBILE_PHONE         + '1'].value = "Mobile_phone"
        outsheet[HOME_ADDRESS         + '1'].value = "Home_address"
        outsheet[HOME_STREET          + '1'].value = "Home_street"
        outsheet[HOME_CITY            + '1'].value = "Home_city"
        outsheet[HOME_STATE           + '1'].value = "Home_state"
        outsheet[HOME_POSTAL_CODE     + '1'].value = "Home_postal_code"
        outsheet[HOME_COUNTRY         + '1'].value = "Home_country"
        outsheet[CONTACT_MAIN_PHONE   + '1'].value = "Contact_main_phone"
        outsheet[BUSINESS_PHONE       + '1'].value = "Business_phone"
        outsheet[BUSINESS_PHONE2      + '1'].value = "Business_phone2"
        outsheet[BUSINESS_FAX         + '1'].value = "Business_fax"
        outsheet[COMPANY              + '1'].value = "Company"
        outsheet[JOB_TITLE            + '1'].value = "Job_title"
        outsheet[DEPARTMENT           + '1'].value = "Department"
        outsheet[OFFICE_LOCATION      + '1'].value = "Office_location"
        outsheet[BUSINESS_ADDRESS     + '1'].value = "Business_address"
        outsheet[BUSINESS_STREET      + '1'].value = "Business_street"
        outsheet[BUSINESS_CITY        + '1'].value = "Business_city"
        outsheet[BUSINESS_STATE       + '1'].value = "Business_state"
        outsheet[BUSINESS_POSTAL_CODE + '1'].value = "Business_postal_code"
        outsheet[BUSINESS_COUNTRY     + '1'].value = "Business_country"
        outsheet[CATEGORIES           + '1'].value = "Categories"
        outsheet[CONNECTED_ON         + '1'].value = "Connected On"
        outsheet[LATITUDE             + '1'].value = "Latitude"
        outsheet[LONGITUDE            + '1'].value = "Longitude"

        dupesheet[FIRST_NAME           + '1'].value = "First_name"
        dupesheet[MIDDLE_NAME          + '1'].value = "Middle_name"
        dupesheet[LAST_NAME            + '1'].value = "Last_name"
        dupesheet[TITLE                + '1'].value = "Title"
        dupesheet[SUFFIX               + '1'].value = "Suffix"
        dupesheet[WEB_PAGE             + '1'].value = "Web_page"
        dupesheet[NOTES                + '1'].value = "Notes"
        dupesheet[EMAIL_ADDRESS        + '1'].value = "Email_address"
        dupesheet[EMAIL_ADDRESS2       + '1'].value = "Email_address2"
        dupesheet[EMAIL_ADDRESS3       + '1'].value = "Email_address3"
        dupesheet[HOME_PHONE           + '1'].value = "Home_phone"
        dupesheet[MOBILE_PHONE         + '1'].value = "Mobile_phone"
        dupesheet[HOME_ADDRESS         + '1'].value = "Home_address"
        dupesheet[HOME_STREET          + '1'].value = "Home_street"
        dupesheet[HOME_CITY            + '1'].value = "Home_city"
        dupesheet[HOME_STATE           + '1'].value = "Home_state"
        dupesheet[HOME_POSTAL_CODE     + '1'].value = "Home_postal_code"
        dupesheet[HOME_COUNTRY         + '1'].value = "Home_country"
        dupesheet[CONTACT_MAIN_PHONE   + '1'].value = "Contact_main_phone"
        dupesheet[BUSINESS_PHONE       + '1'].value = "Business_phone"
        dupesheet[BUSINESS_PHONE2      + '1'].value = "Business_phone2"
        dupesheet[BUSINESS_FAX         + '1'].value = "Business_fax"
        dupesheet[COMPANY              + '1'].value = "Company"
        dupesheet[JOB_TITLE            + '1'].value = "Job_title"
        dupesheet[DEPARTMENT           + '1'].value = "Department"
        dupesheet[OFFICE_LOCATION      + '1'].value = "Office_location"
        dupesheet[BUSINESS_ADDRESS     + '1'].value = "Business_address"
        dupesheet[BUSINESS_STREET      + '1'].value = "Business_street"
        dupesheet[BUSINESS_CITY        + '1'].value = "Business_city"
        dupesheet[BUSINESS_STATE       + '1'].value = "Business_state"
        dupesheet[BUSINESS_POSTAL_CODE + '1'].value = "Business_postal_code"
        dupesheet[BUSINESS_COUNTRY     + '1'].value = "Business_country"
        dupesheet[CATEGORIES           + '1'].value = "Categories"
        dupesheet[CONNECTED_ON         + '1'].value = "Connected On"
        dupesheet[LATITUDE             + '1'].value = "Latitude"
        dupesheet[LONGITUDE            + '1'].value = "Longitude"
    #}}}

    for row in range (first, last + 1):
        # if the previous value is blank we create the new object and store information in it
        '''
        first_name = str(sheet[first_name_col + str(row)].value)
        if first_name != "":
            standardize_str(first_name)
        '''
        compareCriteria = ""
        for col in cols:
            compareCriteria = compareCriteria + standardize_str(sheet[col + str(row)].value) + " "
        if row == first:
            # compare = first_name + " " + last_name
            compare = compareCriteria
            contact = new_contact_from_sheet(sheet, row)
            count = count + 1
        else:
            # current = first_name + " " + last_name
            current = compareCriteria
            match = edit_distance(compare, current, low_threshold, high_threshold)
            matchingSuffixes = standardize_str(sheet[SUFFIX + str(row)].value) == standardize_str(sheet[SUFFIX + str(row - 1)].value)
            # combine information and move on
            if match and matchingSuffixes:
                contact = update_contact_from_sheet(sheet, row, contact)
                '''
                - if there is a duplicate
                - store previous item in list in duplicate list
                - store the duplicate in next spot in duplicate list
                '''
                # combine information
                # keep original
                dupes = dupes + 1

                dupeContact = new_contact_from_sheet(sheet, row - 1)

                #{{{
                dupesheet[FIRST_NAME           + str(dupes)].value = dupeContact.first_name 
                dupesheet[MIDDLE_NAME          + str(dupes)].value = dupeContact.middle_name
                dupesheet[LAST_NAME            + str(dupes)].value = dupeContact.last_name
                dupesheet[TITLE                + str(dupes)].value = dupeContact.title
                dupesheet[SUFFIX               + str(dupes)].value = dupeContact.suffix
                dupesheet[WEB_PAGE             + str(dupes)].value = dupeContact.web_page
                dupesheet[NOTES                + str(dupes)].value = dupeContact.notes
                dupesheet[EMAIL_ADDRESS        + str(dupes)].value = dupeContact.email_address
                dupesheet[EMAIL_ADDRESS2       + str(dupes)].value = dupeContact.email_address2
                dupesheet[EMAIL_ADDRESS3       + str(dupes)].value = dupeContact.email_address3
                dupesheet[HOME_PHONE           + str(dupes)].value = dupeContact.home_phone
                dupesheet[MOBILE_PHONE         + str(dupes)].value = dupeContact.mobile_phone
                dupesheet[HOME_ADDRESS         + str(dupes)].value = dupeContact.home_address
                dupesheet[HOME_STREET          + str(dupes)].value = dupeContact.home_street
                dupesheet[HOME_CITY            + str(dupes)].value = dupeContact.home_city
                dupesheet[HOME_STATE           + str(dupes)].value = dupeContact.home_state
                dupesheet[HOME_POSTAL_CODE     + str(dupes)].value = dupeContact.home_postal_code
                dupesheet[HOME_COUNTRY         + str(dupes)].value = dupeContact.home_country
                dupesheet[CONTACT_MAIN_PHONE   + str(dupes)].value = dupeContact.contact_main_phone
                dupesheet[BUSINESS_PHONE       + str(dupes)].value = dupeContact.business_phone
                dupesheet[BUSINESS_PHONE2      + str(dupes)].value = dupeContact.business_phone2
                dupesheet[BUSINESS_FAX         + str(dupes)].value = dupeContact.business_fax
                dupesheet[COMPANY              + str(dupes)].value = dupeContact.company
                dupesheet[JOB_TITLE            + str(dupes)].value = dupeContact.job_title
                dupesheet[DEPARTMENT           + str(dupes)].value = dupeContact.department
                dupesheet[OFFICE_LOCATION      + str(dupes)].value = dupeContact.office_location
                dupesheet[BUSINESS_ADDRESS     + str(dupes)].value = dupeContact.business_address
                dupesheet[BUSINESS_STREET      + str(dupes)].value = dupeContact.business_street
                dupesheet[BUSINESS_CITY        + str(dupes)].value = dupeContact.business_city
                dupesheet[BUSINESS_STATE       + str(dupes)].value = dupeContact.business_state
                dupesheet[BUSINESS_POSTAL_CODE + str(dupes)].value = dupeContact.business_postal_code
                dupesheet[BUSINESS_COUNTRY     + str(dupes)].value = dupeContact.business_country
                dupesheet[CATEGORIES           + str(dupes)].value = dupeContact.categories
                dupesheet[CONNECTED_ON         + str(dupes)].value = dupeContact.connected_on
                dupesheet['AI' + str(dupes)].value = row - 1
                #}}}

                # keep duplicate
                dupes = dupes + 1

                dupeContact = new_contact_from_sheet(sheet, row)
                #{{{

                dupesheet[FIRST_NAME           + str(dupes)].value = dupeContact.first_name 
                dupesheet[MIDDLE_NAME          + str(dupes)].value = dupeContact.middle_name
                dupesheet[LAST_NAME            + str(dupes)].value = dupeContact.last_name
                dupesheet[TITLE                + str(dupes)].value = dupeContact.title
                dupesheet[SUFFIX               + str(dupes)].value = dupeContact.suffix
                dupesheet[WEB_PAGE             + str(dupes)].value = dupeContact.web_page
                dupesheet[NOTES                + str(dupes)].value = dupeContact.notes
                dupesheet[EMAIL_ADDRESS        + str(dupes)].value = dupeContact.email_address
                dupesheet[EMAIL_ADDRESS2       + str(dupes)].value = dupeContact.email_address2
                dupesheet[EMAIL_ADDRESS3       + str(dupes)].value = dupeContact.email_address3
                dupesheet[HOME_PHONE           + str(dupes)].value = dupeContact.home_phone
                dupesheet[MOBILE_PHONE         + str(dupes)].value = dupeContact.mobile_phone
                dupesheet[HOME_ADDRESS         + str(dupes)].value = dupeContact.home_address
                dupesheet[HOME_STREET          + str(dupes)].value = dupeContact.home_street
                dupesheet[HOME_CITY            + str(dupes)].value = dupeContact.home_city
                dupesheet[HOME_STATE           + str(dupes)].value = dupeContact.home_state
                dupesheet[HOME_POSTAL_CODE     + str(dupes)].value = dupeContact.home_postal_code
                dupesheet[HOME_COUNTRY         + str(dupes)].value = dupeContact.home_country
                dupesheet[CONTACT_MAIN_PHONE   + str(dupes)].value = dupeContact.contact_main_phone
                dupesheet[BUSINESS_PHONE       + str(dupes)].value = dupeContact.business_phone
                dupesheet[BUSINESS_PHONE2      + str(dupes)].value = dupeContact.business_phone2
                dupesheet[BUSINESS_FAX         + str(dupes)].value = dupeContact.business_fax
                dupesheet[COMPANY              + str(dupes)].value = dupeContact.company
                dupesheet[JOB_TITLE            + str(dupes)].value = dupeContact.job_title
                dupesheet[DEPARTMENT           + str(dupes)].value = dupeContact.department
                dupesheet[OFFICE_LOCATION      + str(dupes)].value = dupeContact.office_location
                dupesheet[BUSINESS_ADDRESS     + str(dupes)].value = dupeContact.business_address
                dupesheet[BUSINESS_STREET      + str(dupes)].value = dupeContact.business_street
                dupesheet[BUSINESS_CITY        + str(dupes)].value = dupeContact.business_city
                dupesheet[BUSINESS_STATE       + str(dupes)].value = dupeContact.business_state
                dupesheet[BUSINESS_POSTAL_CODE + str(dupes)].value = dupeContact.business_postal_code
                dupesheet[BUSINESS_COUNTRY     + str(dupes)].value = dupeContact.business_country
                dupesheet[CATEGORIES           + str(dupes)].value = dupeContact.categories
                dupesheet[CONNECTED_ON         + str(dupes)].value = dupeContact.connected_on
                dupesheet['AI' + str(dupes)].value = row
                #}}}

                dupes = dupes + 1

                # save the combined contact
                #{{{
                dupesheet[FIRST_NAME           + str(dupes)].value = contact.first_name 
                dupesheet[MIDDLE_NAME          + str(dupes)].value = contact.middle_name
                dupesheet[LAST_NAME            + str(dupes)].value = contact.last_name
                dupesheet[TITLE                + str(dupes)].value = contact.title
                dupesheet[SUFFIX               + str(dupes)].value = contact.suffix
                dupesheet[WEB_PAGE             + str(dupes)].value = contact.web_page
                dupesheet[NOTES                + str(dupes)].value = contact.notes
                dupesheet[EMAIL_ADDRESS        + str(dupes)].value = contact.email_address
                dupesheet[EMAIL_ADDRESS2       + str(dupes)].value = contact.email_address2
                dupesheet[EMAIL_ADDRESS3       + str(dupes)].value = contact.email_address3
                dupesheet[HOME_PHONE           + str(dupes)].value = contact.home_phone
                dupesheet[MOBILE_PHONE         + str(dupes)].value = contact.mobile_phone
                dupesheet[HOME_ADDRESS         + str(dupes)].value = contact.home_address
                dupesheet[HOME_STREET          + str(dupes)].value = contact.home_street
                dupesheet[HOME_CITY            + str(dupes)].value = contact.home_city
                dupesheet[HOME_STATE           + str(dupes)].value = contact.home_state
                dupesheet[HOME_POSTAL_CODE     + str(dupes)].value = contact.home_postal_code
                dupesheet[HOME_COUNTRY         + str(dupes)].value = contact.home_country
                dupesheet[CONTACT_MAIN_PHONE   + str(dupes)].value = contact.contact_main_phone
                dupesheet[BUSINESS_PHONE       + str(dupes)].value = contact.business_phone
                dupesheet[BUSINESS_PHONE2      + str(dupes)].value = contact.business_phone2
                dupesheet[BUSINESS_FAX         + str(dupes)].value = contact.business_fax
                dupesheet[COMPANY              + str(dupes)].value = contact.company
                dupesheet[JOB_TITLE            + str(dupes)].value = contact.job_title
                dupesheet[DEPARTMENT           + str(dupes)].value = contact.department
                dupesheet[OFFICE_LOCATION      + str(dupes)].value = contact.office_location
                dupesheet[BUSINESS_ADDRESS     + str(dupes)].value = contact.business_address
                dupesheet[BUSINESS_STREET      + str(dupes)].value = contact.business_street
                dupesheet[BUSINESS_CITY        + str(dupes)].value = contact.business_city
                dupesheet[BUSINESS_STATE       + str(dupes)].value = contact.business_state
                dupesheet[BUSINESS_POSTAL_CODE + str(dupes)].value = contact.business_postal_code
                dupesheet[BUSINESS_COUNTRY     + str(dupes)].value = contact.business_country
                dupesheet[CATEGORIES           + str(dupes)].value = contact.categories
                dupesheet[CONNECTED_ON         + str(dupes)].value = contact.connected_on
                #}}}

                # create a blank space
                dupes = dupes + 1
            # store the information and create a new contact
            else:
                # store information
                #{{{
                outsheet[FIRST_NAME           + str(count)].value = contact.first_name 
                outsheet[MIDDLE_NAME          + str(count)].value = contact.middle_name
                outsheet[LAST_NAME            + str(count)].value = contact.last_name
                outsheet[TITLE                + str(count)].value = contact.title
                outsheet[SUFFIX               + str(count)].value = contact.suffix
                outsheet[WEB_PAGE             + str(count)].value = contact.web_page
                outsheet[NOTES                + str(count)].value = contact.notes
                outsheet[EMAIL_ADDRESS        + str(count)].value = contact.email_address
                outsheet[EMAIL_ADDRESS2       + str(count)].value = contact.email_address2
                outsheet[EMAIL_ADDRESS3       + str(count)].value = contact.email_address3
                outsheet[HOME_PHONE           + str(count)].value = contact.home_phone
                outsheet[MOBILE_PHONE         + str(count)].value = contact.mobile_phone
                outsheet[HOME_ADDRESS         + str(count)].value = contact.home_address
                outsheet[HOME_STREET          + str(count)].value = contact.home_street
                outsheet[HOME_CITY            + str(count)].value = contact.home_city
                outsheet[HOME_STATE           + str(count)].value = contact.home_state
                outsheet[HOME_POSTAL_CODE     + str(count)].value = contact.home_postal_code
                outsheet[HOME_COUNTRY         + str(count)].value = contact.home_country
                outsheet[CONTACT_MAIN_PHONE   + str(count)].value = contact.contact_main_phone
                outsheet[BUSINESS_PHONE       + str(count)].value = contact.business_phone
                outsheet[BUSINESS_PHONE2      + str(count)].value = contact.business_phone2
                outsheet[BUSINESS_FAX         + str(count)].value = contact.business_fax
                outsheet[COMPANY              + str(count)].value = contact.company
                outsheet[JOB_TITLE            + str(count)].value = contact.job_title
                outsheet[DEPARTMENT           + str(count)].value = contact.department
                outsheet[OFFICE_LOCATION      + str(count)].value = contact.office_location
                outsheet[BUSINESS_ADDRESS     + str(count)].value = contact.business_address
                outsheet[BUSINESS_STREET      + str(count)].value = contact.business_street
                outsheet[BUSINESS_CITY        + str(count)].value = contact.business_city
                outsheet[BUSINESS_STATE       + str(count)].value = contact.business_state
                outsheet[BUSINESS_POSTAL_CODE + str(count)].value = contact.business_postal_code
                outsheet[BUSINESS_COUNTRY     + str(count)].value = contact.business_country
                outsheet[CATEGORIES           + str(count)].value = contact.categories
                outsheet[CONNECTED_ON         + str(count)].value = contact.connected_on
                outsheet[LATITUDE             + str(count)].value = contact.latitude
                outsheet[LONGITUDE            + str(count)].value = contact.longitude
                #}}}
                # reset compare value
                # compare = first_name + " " + last_name
                compare = compareCriteria
                # create a new contact
                contact = new_contact_from_sheet(sheet, row)
                count = count + 1

    if printing:
        print()
        print("Out of " + str(1 + last - first) + " companies " + str(count) + " were unique contacts")
        print("Saving...")

    out.save("purged" + str(cols) + ".xlsx")
    dupe.save("duplicates" + str(cols) + ".xlsx")

    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(3)
    pygame.mixer.music.stop()

    if printing:
        print()
        print("Done!")
#}}}

# Fix Country Names
# fix_country_names(sheet)
# {{{
def fix_country_names(fileName, sheetName):
    if printing:
        print("Opening...")
    wb = openpyxl.load_workbook(fileName + ".xlsx")
    sheet = wb[sheetName]

    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')

    compare = ""
    current = ""
    count = 1
    first = 2
    last = sheet.max_row

    for row in range (first, last + 1):
        # if the previous value is blank we create the new object and store information in it
        country = str(sheet[BUSINESS_COUNTRY + str(row)].value)
        if country == 'None':
            break
        if row == first:
            compare = country
            count = count + 1
        else:
            current = country
            match = edit_distance(compare, current, low_threshold, high_threshold)
            # combine information and move on
            if match:
                if printing:
                    print("Now changing row " + str(row))
                sheet[BUSINESS_COUNTRY + str(row)].value = compare
            else:
                # reset compare value
                compare = country
                count = count + 1

    if printing:
        print("Saving...")
    wb.save("newCombined.xlsx")

    pygame.mixer.music.play()
    time.sleep(3)
    pygame.mixer.music.stop()
#}}}

# combine_contacts(2, 100)
# combine_contacts("gmail", "contacts", 'linkedin', 'contacts')
# remove_duplicate_contacts("combined", "contacts", 2, 6903)
# combine_with_CMA("CMAShipping", "Attendees", "gli", "contacts")
# standardize_USA(sys.argv)
remove_duplicate_contacts(sys.argv)
