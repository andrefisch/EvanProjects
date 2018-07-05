
# coding: utf-8

# In[15]:

from urllib.request import Request, urlopen
from datetime import *
from bs4 import BeautifulSoup as bs
import openpyxl
import pygame
import re
import time


# In[16]:

url = 'https://www.wista.net/en/wistabook/wistabook/wista_lang,33'

request = Request(url)
json = urlopen(request).read()
array = str(json)
soup = bs(array)
prettyHTML = soup.prettify()
prettyHTML = prettyHTML.split('\n')


# In[ ]:

urls = []
regexURL = 'href="(.*?)"'
for i in range(977, len(prettyHTML)):
    if 'show_profile' in prettyHTML[i]:
        matchURL = re.search(regexURL, prettyHTML[i])
        if matchURL:
            urls.append(matchURL.group(1))


# In[ ]:

dicty = {}
dicty['PROFILE']                = "A"
dicty['COMPANY']                = "B"
dicty['TITLE']                  = "C"
dicty['ADDRESS']                = "D"
dicty['EMAIL']                  = "E"
dicty['WEBSITE']                = "F"
dicty['FUNCTIONALCATEGORIES']   = "G"
dicty['SHIPPINGCATEGORIES']     = "H"
dicty['JOBDESCRIPTION']         = "I"
dicty['DESCRIPTIONOFCOMPANY']   = "J"
dicty['STARTEDINSHIPPING']      = "K"
dicty['OFFICEPHONE']            = 'L'
dicty['MOBILEPHONE']            = 'M'
dicty['FAX']                    = 'N'
dicty['EXPERIENCE']             = 'O'
dicty['LASTUPDATE']             = "P"

# Create a new excel file
out = openpyxl.Workbook()
# Open the worksheet we want to edit
outsheet = out.create_sheet("wista members")
# if 'sheet' appears randomly we can delete it
rm = out.get_sheet_by_name('Sheet')
out.remove_sheet(rm)
#################
# DO STUFF HERE #
#################
outsheet[dicty['PROFILE']               + '1'].value = "Profile"
outsheet[dicty['COMPANY']               + '1'].value = "Company"
outsheet[dicty['TITLE']                 + '1'].value = "Title"
outsheet[dicty['ADDRESS']               + '1'].value = "Address"
outsheet[dicty['EMAIL']                 + '1'].value = "Email"
outsheet[dicty['WEBSITE']               + '1'].value = "Website"
outsheet[dicty['FUNCTIONALCATEGORIES']  + '1'].value = "Functioning Categories"
outsheet[dicty['SHIPPINGCATEGORIES']    + '1'].value = "Shipping Categories"
outsheet[dicty['JOBDESCRIPTION']        + '1'].value = "Job Description"
outsheet[dicty['DESCRIPTIONOFCOMPANY']  + '1'].value = "Company Description"
outsheet[dicty['STARTEDINSHIPPING']     + '1'].value = "Started"
outsheet[dicty['OFFICEPHONE']           + '1'].value = 'Office Phone'
outsheet[dicty['MOBILEPHONE']           + '1'].value = 'Mobile Phone'
outsheet[dicty['FAX']                   + '1'].value = 'Fax'
outsheet[dicty['EXPERIENCE']            + '1'].value = 'Experience'
outsheet[dicty['LASTUPDATE']            + '1'].value = "Last Updated"

regexSpaces = '^ *'

row = 2
step = 50
for j in range(0, len(urls)):
    if row % step == 0:
        print('TAKING A BREAK TO SAVE')
        out.save("wista.xlsx")
    request = Request(urls[j])
    print(urls[j])
    json = urlopen(request).read()
    array = str(json)
    soup = bs(array)
    prettyHTML = soup.prettify()
    prettyHTML = prettyHTML.split('\n')
    
    info = []
    for i in range(630, len(prettyHTML)):
        line = prettyHTML[i]
        line = re.sub(regexSpaces, '', line)
        if '<' not in line and '\\n' not in line and len(info) < 100:
            # skip experience, it is often blank and screws up everything else
            info.append(line)
            print(line)
    
    if 'Error' in info[0]:
        continue
    
    current = ''
    newInfo = []
    after = 0
    # if line is a keyword
    # - add line to list
    # - start recording
    # if we encounter 'Last update' we are done
    print(info)
    for i in range(0, len(info)):
        if info[i].upper().replace(' ', '').replace('-', '').strip() in dicty:
            if current != "":
                newInfo.append(current)
                print('adding info:', current)
            newInfo.append(info[i])
            print('adding keyword:', info[i])
            current = ""
            if ('Last update') in info[i]:
                newInfo.append(info[i + 1])
        else:
            current = current + ' ' + info[i]

    # if two adjacent values are keywords remove the first one
    # endless loop
    # - go through list
    # - if two adjacent values are both dict keys, remove first one
    # - log that a change was made, break
    last = False
    print(newInfo)
    while(True):
        change = False
        for i in range(0, len(newInfo)):
            if last and newInfo[i].upper().replace(' ', '').replace('-', '').strip() in dicty:
                del newInfo[i - 1]
                change = True
                last = False
                break
            else:
                last = False
            if newInfo[i].upper().replace(' ', '').replace('-', '').strip() in dicty:
                last = True
            else:
                last = False
        if not change:
            break

    print(newInfo)
    for i in range(0, len(newInfo), 2):
        # if newInfo[i + 1].upper().replace(' ', '').replace('-', '') == 'LASTUPDATE':
            # break
        column = newInfo[i].upper().replace(' ', '').replace('-', '')
        inf = newInfo[i + 1]
        print(row, column, inf)
        outsheet[dicty[column] + str(row)].value = inf
        if column == 'LASTUPDATE':
            print('breaking:', newInfo[i], newInfo[i + 1])
            break
    row = row + 1
        

# Save the file
out.save("wista.xlsx")

# LMK when the script is done
pygame.init()
pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
pygame.mixer.music.play()
time.sleep(5)
pygame.mixer.music.stop()


# In[ ]:



