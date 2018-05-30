import openpyxl
import pygame
import re
import string
import time

printing = True
saving = True
# formatting phone numbers

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

# Remove all spaces from the last n nonspace characters of the given string
# remove_end_space(string, num)
#{{{
def remove_end_space(string, num):
    string2 = string[::-1]
    count = 0
    index = 0
    for i in range(0, len(string2)):
        if string2[i] != " ":
            count += 1
        if count >= num:
            index = i
            break
    first = string2[index:]
    second = string2[:index]
    second = second.replace(" ", "")
    return (second + first)[::-1]
#}}}

# Extract first, last, and middle name from a name with more than 3 parts
# determine_names(listy)
#{{{
def determine_names(listy):
    dicty   = {}
    lasty   = []
    middley = []
    # first spot is always first name at this point
    dicty['first_name'] = listy[0]
    del listy[0]
    # - reverse list 
    # - take first name in reversed list (last name) and add it to last name list, delete it
    # - look at next name and see if it is capitalized
    #   - if not add to last name list, repeat
    #   - otherwise add this and rest to middle name list
    listy = listy[::-1]
    lasty.append(listy[0])
    del listy[0]
    lasts = True
    for i in range(0, len(listy)):
        if (not listy[i].istitle()) and lasts:
            lasty.insert(0, listy[i])
        else:
            lasts = False
            middley.insert(0, listy[i])

    dicty['middle_name'] = ' '.join(middley)
    dicty['last_name'] = ' '.join(lasty)
    return dicty
#}}}

# formatting_phone_number(number)
#{{{
def formatting_name(name):
    original = name
    names = {"first_name": "", "middle_name": "", "last_name": "", "appellation": ""}

    # remove unnecessary suffixes
    regexSuffix = ',? (Jr\.? ?|Sr\.? ?|Ph\.? ?D\.? ?|P\.?Ehj)'
    name = re.sub(regexSuffix, "", name, re.IGNORECASE)

    # extract appellation 
    regexAppellation = '^(Mr\.?|Mrs\.?|Ms\.?|Rev\.?|Hon\.?|Dr\.?|Capt\.?|Dcn\.?|Amb\.?|Lt\.?|MIDN\.?|Miss\.?|Fr\.?) (.*)'
    # number = re.sub(regexAppellation, "", number)
    # regex0 = '^0+(.*)'
    matchAppellation = re.search(regexAppellation, name, re.IGNORECASE)
    if matchAppellation:
        names['appellation'] = matchAppellation.group(1)
        name = matchAppellation.group(2)

    # split names
    matchList = []
    regexName = "([\w+\.-]+)"
    while True:
        matchName = re.search(regexName, name)
        if matchName:
            matchList.append(matchName.group(1))
            name = name.replace(matchName.group(1), "")
        else:
            break

    # if there are only one, two, or three names in the list it is easy
    nameLen = len(matchList)
    if nameLen == 1:
        names['last_name'] = matchList[0]
    elif nameLen == 2:
        names['first_name'] = matchList[0]
        names['last_name'] = matchList[-1]
    elif nameLen == 3:
        names['first_name'] = matchList[0]
        names['middle_name'] = matchList[1]
        names['last_name'] = matchList[-1]
    else:
        names = {**names, **determine_names(matchList)}

    '''
    if printing:
        if original != "None":
            print(original + " -> " + name)
    '''

    return names
#}}}

# Format all numbers in a column
# format_all_numbers(fileName, sheetName, col)
#{{{
def format_all_numbers(fileName, sheetName, *cols):
    if printing:
        print("Opening...")
    wb = openpyxl.load_workbook(fileName + ".xlsx")
    sheet = wb[sheetName]

    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')

    first = 2
    last = sheet.max_row

    for col in cols:
        for row in range (first, last + 1):
            number = str(sheet[col + str(row)].value)
            formatted = formatting_phone_number(number)
            if saving:
                sheet[col + str(row)].value = formatted

    if printing:
        print("Processing " + str(last - first) + " rows...")

    if printing and saving:
        print("Saving...")
        wb.save("betterNumbers.xlsx")

    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()
#}}}

# test stuff
#{{{
name1 = "Mr. Evangelos Efstathiou, Ph.D."
name2 = "Mr. Bradley D.M. Golden"
name3 = "Mr. Jan-Willem Ovind van den Dijssel"

print(formatting_name(name1))
print(formatting_name(name2))
print(formatting_name(name3))
#}}}
