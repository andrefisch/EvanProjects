import math
import openpyxl
import pygame, time
import re
import sys

pattern = 'href="/projects/development-projects/([-1-9a-z])'

# Open the file for editing
wb = openpyxl.Workbook()
# Open the worksheet we want to edit
sheet = wb.create_sheet("Emails")
# Open the finished playing sound
pygame.init()
ygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')

print(wb.get_sheet_names())

rm = wb.get_sheet_by_name('Sheet')
wb.remove_sheet(rm)

saving = True
printing = True

if saving:
    sheet['A1'] = "Name"
    sheet['B1'] = "Email Address"
    sheet['C1'] = "Domain Name"

string = ""
table = {}
with open(sys.argv[1]) as f:
    for line in f:
        string = line
line = re.sub('["\']', '', line.lower())
emails = line.split(";")
row = 1
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
        row = row + 1
        if saving:
            table[email] = [name, domain]
            
row = 0
for key, value in sorted(table.items(), key=lambda e: e[1][1]):
    row = row + 1
    email  = key
    name   = value[0]
    ind = name.find('@')
    if ind > -1 or name == "":
        name = email
        ind = name.find('@')
        name = re.sub('[\._]', ' ', name[0:ind])
    name = re.sub('[0-9]', '', name.title())
    domain = value[1]
    if printing:
        print(name, email, domain)
    sheet['A' + str(row)] = name
    sheet['B' + str(row)] = email
    sheet['C' + str(row)] = domain

wb.save("emails.xlsx")
pygame.mixer.music.play()
time.sleep(3)
pygame.mixer.music.stop()
