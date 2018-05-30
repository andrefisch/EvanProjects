import math
import openpyxl
import pygame, time
import re
import sys
from splitNames import determine_names

pattern = 'href="/projects/development-projects/([-1-9a-z])'

FIRST_NAME    = "A"
MIDDLE_NAME   = "B"
LAST_NAME     = "C"
EMAIL_ADDRESS = "D"
DOMAIN_NAME   = "E"

# Open the file for editing
wb = openpyxl.Workbook()
# Open the worksheet we want to edit
sheet = wb.create_sheet("Emails")
# Open the finished playing sound
pygame.init()
pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')

print(wb.get_sheet_names())

rm = wb.get_sheet_by_name('Sheet')
wb.remove_sheet(rm)

saving = True
printing = True

if saving:
    sheet[FIRST_NAME    + '1'] = "First Name"
    sheet[MIDDLE_NAME   + '1'] = "Middle Name"
    sheet[LAST_NAME     + '1'] = "Last Name"
    sheet[EMAIL_ADDRESS + '1'] = "Email Address"
    sheet[DOMAIN_NAME   + '1'] = "Domain Name"

string = ""
table = {}
with open(sys.argv[1]) as f:
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
    sheet[FIRST_NAME    + str(row)] = dicty['first_name'].title().strip()
    sheet[MIDDLE_NAME   + str(row)] = dicty['middle_name'].title().strip()
    sheet[LAST_NAME     + str(row)] = dicty['last_name'].title().strip()
    sheet[EMAIL_ADDRESS + str(row)] = email
    sheet[DOMAIN_NAME   + str(row)] = domain

wb.save("emails.xlsx")
pygame.mixer.music.play()
time.sleep(5)
pygame.mixer.music.stop()
