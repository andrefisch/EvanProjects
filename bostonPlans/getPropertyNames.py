from urllib.parse import urlencode
from urllib.request import Request, urlopen
import pandas
import openpyxl
import re
import math
import pygame, time

begin = 44

url = 'http://www.bostonplans.org/projects/development-projects/?viewall=1'
request = Request(url)
json = urlopen(request).read()
array = str(json)[2:-1]
array = str(array).split("\\n")

pattern = 'href="/projects/development-projects/([-1-9a-z])'
#listy = re.findall(pattern, array)

smallArray = []
for i in range (0, len(array)):
    if 'href="/projects/development-projects/' in array[i]:
        temp = array[i][begin:]
        stop = temp.find('"')
        smallArray.append(temp[:stop])
        print(temp[:stop])

# Open the file for editing
wb = openpyxl.Workbook()
# Open the worksheet we want to edit
sheet = wb.create_sheet("Properties")
# Open the finished playing sound
pygame.init()
pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')

baseUrl = 'http://www.bostonplans.org/projects/development-projects/'
saving = True
printing = False

if saving:
    sheet['A1'] = "Project Name"
    sheet['B1'] = "Neighborhood"
    sheet['C1'] = "Address"
    sheet['D1'] = "Land Sq. Feet"
    sheet['E1'] = "Building Size"
    sheet['F1'] = "Email Address"
    sheet['G1'] = "Project Manager"
    sheet['H1'] = "Uses"
    sheet['I1'] = "Residential Units"
    sheet['J1'] = "Project Description"

start = 1
end = len(smallArray)
for i in range(start, end):
    print (str(format((row - start) / (end - start) * 100.00, '.2f')) + "%: " + smallArray[i]
    if saving: sheet['A' + str(i + 1)] = smallArray[i]
    request = Request(baseUrl + smallArray[i])
    json = urlopen(request).read()
    projectInfo = str(json)[2:-1]
    projectInfo = str(projectInfo).split("\\n")
    #print(projectInfo)
    relevantInfo = []
    neighborhood, address, sqFeet, buildingSize, projectManager, emailAddress, projectDescription = "", "", "", "", "", "", ""
    restart = 0
    fork = 0
    for j in range(restart, len(projectInfo)):
        if ">Neighborhood<" in projectInfo[j]:
            fork = 0
            restart = j + 1
            neighborhood = projectInfo[j][85:-6]
            if saving: sheet['B' + str(i + 1)] = neighborhood
            if printing: print (neighborhood)
            break
        elif ">Neighborhood:<" in projectInfo[j]:
            fork = 1
            restart = 0
            neighborhood = projectInfo[j][35:-9]
            if saving: sheet['B' + str(i + 1)] = neighborhood
            if printing: print (neighborhood)
            break
    if fork == 0:
        for j in range(restart, len(projectInfo)):
            if ">Address<" in projectInfo[j]:
                restart = j + 1
                address = projectInfo[j][180:-10]
                if saving: sheet['C' + str(i + 1)] = address
                if printing: print(address)
                break
        for j in range(restart, len(projectInfo)):
            if ">Land Sq. Feet<" in projectInfo[j]:
                restart = j + 1
                sqFeet = projectInfo[j][86:-6]
                if saving: sheet['D' + str(i + 1)] = sqFeet
                if printing: print(sqFeet)
                break
        for j in range(restart, len(projectInfo)):
            if ">Building Size<" in projectInfo[j]:
                restart = j + 1
                buildingSize = projectInfo[j][86:-6]
                if saving: sheet['E' + str(i + 1)] = buildingSize
                if printing: print(buildingSize)
                break
        for j in range(restart, len(projectInfo)):
            if ">Project Manager<" in projectInfo[j]:
                restart = j + 1
                # get email address too
                smallStr = projectInfo[j + 1][73:-11]
                split = smallStr.find('\\')
                emailAddress = smallStr[:split]
                if saving: sheet['F' + str(i + 1)] = emailAddress
                if printing: print(emailAddress)
                projectManager = smallStr[split + 4:]
                if saving: sheet['G' + str(i + 1)] = projectManager
                if printing: print(projectManager)
                break
        for j in range(restart, len(projectInfo)):
            if ">Project Description<" in projectInfo[j]:
                projectDescription = projectInfo[j + 1][37:-6]
                if saving: sheet['J' + str(i + 1)] = projectDescription
                if printing: print (projectDescription)
                break
    if fork == 1:
        for j in range(restart, len(projectInfo)):
            if ">Address:<" in projectInfo[j]:
                restart = j + 1
                address = projectInfo[j][30:-9]
                if saving: sheet['C' + str(i + 1)] = address
                if printing: print(address)
                break
        for j in range(restart, len(projectInfo)):
            if ">Land Sq. Feet:<" in projectInfo[j]:
                restart = j + 1
                sqFeet = projectInfo[j][36:-9]
                if saving: sheet['D' + str(i + 1)] = sqFeet
                if printing: print(sqFeet)
                break
        for j in range(restart, len(projectInfo)):
            if ">Building Size:<" in projectInfo[j]:
                restart = j + 1
                buildingSize = projectInfo[j][36:-9]
                if saving: sheet['E' + str(i + 1)] = buildingSize
                if printing: print(buildingSize)
                break
        for j in range(restart, len(projectInfo)):
            if ">Uses:<" in projectInfo[j]:
                restart = j + 1
                # get email address too
                uses = projectInfo[j][27:-9]
                if saving: sheet['H' + str(i + 1)] = uses
                if printing: print(uses)
                break
        for j in range(restart, len(projectInfo)):
            if ">Residential Units:<" in projectInfo[j]:
                units = projectInfo[j][40:-9]
                if saving: sheet['I' + str(i + 1)] = units
                if printing: print (units)
                break
        for j in range(restart, len(projectInfo)):
            if ">Project Description:<" in projectInfo[j]:
                projectDescription = projectInfo[j + 2][4:-2]
                if saving: sheet['J' + str(i + 1)] = projectDescription
                if printing: print (projectDescription)
                break
            
wb.save("properties.xlsx")
pygame.mixer.music.play()
time.sleep(3)
pygame.mixer.music.stop()


# In[ ]:



