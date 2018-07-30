import datetime
import math
import openpyxl
import pygame, time
import re
import sys
from splitNames import determine_names

def emailsToExcel():
    if len(sys.argv) < 3:
        print("copyEmails requires two arguments: the file to process and the name of the list")
        print("copyEmails should be called like this:")
        print("python3 copyEmails.py emails.txt 'RailList'")
        return

    filename = sys.argv[1]
    group = sys.argv[2]

    FIRST_NAME       = "A"
    MIDDLE_NAME      = "B"
    LAST_NAME        = "C"
    EMAIL_ADDRESS    = "D"
    DOMAIN_NAME      = "E"
    GROUP_MEMBERSHIP = "F"
    DATE             = "G"

    # Open the file for editing
    wb = openpyxl.Workbook()
    # Open the worksheet we want to edit
    sheet = wb.create_sheet("Emails")

    rm = wb['Sheet']
    wb.remove(rm)

    saving = True
    printing = True

    if saving:
        sheet[FIRST_NAME       + '1'] = "Given Name"
        sheet[MIDDLE_NAME      + '1'] = "Additional Name"
        sheet[LAST_NAME        + '1'] = "Family Name"
        sheet[EMAIL_ADDRESS    + '1'] = "E-mail 1 - Value"
        sheet[DOMAIN_NAME      + '1'] = "Organization 1 - Name"
        sheet[GROUP_MEMBERSHIP + '1'] = "Group Membership"
        sheet[DATE             + '1'] = "Date"

    NOTCOMPANIES = ['aol', 'comcast', 'gmail', 'hotmail', 'verizon', 'yahoo']

    string = ""
    table = {}
    with open(filename) as f:
        for line in f:
            string = line
    line = re.sub('["\']', '', line.lower())
    emails = line.split(";")
    for i in range(0, len(emails)):
        print (str(format((i) / len(emails) * 100.00, '.2f')) + "%: " + emails[i])
        name, email, domain = "", "", ""
        pattern = '(.*)[(<](.*)[>)]'
        m = re.search(pattern, emails[i])
        # work is done here
        if (m and m.group(2)):
            name = m.group(1)
            email = m.group(2)
            if "," in m.group(1):
                pattern = "(.*), (.*)"
                n = re.search(pattern, m.group(1))
                name = n.group(2) + n.group(1)
        else:
            name = ""
            email = emails[i]
        name = name.strip().title()
        email = email.strip()
        pattern = '@(.*)'
        m = re.search(pattern, email)
        if (not m):
            email = ""
        else:
            domain = m.group(1)
        # saving is done here as long as there is at least one field to save 
        if email != "":
            if saving:
                table[email] = [name, domain]
    row = 1
    for key, value in sorted(table.items(), key=lambda e: e[1][1]):
        row = row + 1
        email  = key
        name   = value[0]
        ind = name.find('@')
        if ind > -1 or name == "":
            name = email
            ind = name.find('@')
            name = re.sub('[\._]', ' ', name[0:ind])
        regexAppellation = '^(Mr\.?|Mrs\.?|Ms\.?|Rev\.?|Hon\.?|Dr\.?|Capt\.?|Dcn\.?|Amb\.?|Lt\.?|MIDN\.?|Miss\.?|Fr\.?) (.*)'
        # number = re.sub(regexAppellation, "", number)
        # regex0 = '^0+(.*)'
        matchAppellation = re.search(regexAppellation, name, re.IGNORECASE)
        if matchAppellation:
            name = matchAppellation.group(2)
        dicty = determine_names(name.split(" "))
        name = re.sub('[0-9]', '', name.title())
        domain = value[1]
        index = domain[::-1].find('.')
        if index >= 0:
            domain = domain[:-index - 1]
        '''
        if printing:
        print(name, email, domain)
        print(dicty)
        '''
        sheet[FIRST_NAME       + str(row)] = dicty['first_name'].title().strip()
        sheet[MIDDLE_NAME      + str(row)] = dicty['middle_name'].title().strip()
        sheet[LAST_NAME        + str(row)] = dicty['last_name'].title().strip()
        sheet[EMAIL_ADDRESS    + str(row)] = email
        if domain not in NOTCOMPANIES:
            sheet[DOMAIN_NAME      + str(row)] = domain
        sheet[GROUP_MEMBERSHIP + str(row)] = group
        sheet[DATE             + str(row)] = datetime.datetime.today().strftime('%m/%d/%Y')

    print(len(emails), "emails processed")

    wb.save("emails.xlsx")

    # Open the finished playing sound
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

emailsToExcel()
