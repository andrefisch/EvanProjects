import openpyxl
import pygame
import re
import string
import sys
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

# Format a single phone number
# formatting_phone_number(number)
#{{{
def formatting_phone_number(number):
    original = number
    # replace certain punctuations with a space
    regexPunct = '([\.:;|/@]| {2,})'
    number = re.sub(regexPunct, " ", number)
    # replace all +'s with nothing
    number = re.sub('[\+\[\]\'=]', "", number)
    # remove all letters
    regexLetters = '[a-zA-Z]'
    number = re.sub(regexLetters, "", number)
    # remove all spaces before the first number
    regexBegin = '^ *'
    number = re.sub(regexBegin, "", number)
    # remove everything after the last number
    regexEnd = '\D*$'
    number = re.sub(regexEnd, "", number)
    # 077, 078 -> +4477 or +4478
    regex077 = '^\D*077(.*)'
    regex078 = '^\D*078(.*)'
    match077 = re.search(regex077, number)
    match078 = re.search(regex078, number)
    # remove leading 001 and 011
    regex01 = '^(001 ?-?|011 ?-?)'
    number = re.sub(regex01, "", number)
    # remove all leading 0's
    regex0 = '^0+(.*)'
    match0 = re.search(regex0, number)

    # if the number starts with 00 replace 00 with +
    # WORKS
    if match0:
        number = "+" + match0.group(1)
    # if the number starts with 077 replace 077 with +4477
    # WORKS
    if match077:
        number = "+4477" + match077.group(1)
    # if the number starts with 078 replace 077 with +4478
    # WORKS
    if match078:
        number = "+4478" + match078.group(1)
    # if number contains a useless (0), remove it
    regexParen = ' ?\([+ 0]?\) ?'
    number = re.sub(regexParen, '', number)
    # if number contains a weird (), remove it
    regex1      = '^\D*1'
    regexUSA    = '^\D*(\+?1?)?\D*(\d{3})\D*(\d{3})\D*(\d{4})$'
    regexUSAno1 = '^(\(\d{3}\))\D*(\d{3})\D*(\d{4})'
    # regexEurope = '\D*(\+?\d{2,3})\D*'
    matchUSA    = re.search(regexUSA,    number)
    matchUSAno1 = re.search(regexUSAno1, number)
    if matchUSA:
        number = "+1 (" + matchUSA.group(2) + ") " + matchUSA.group(3) + "-" + matchUSA.group(4)
    elif matchUSAno1:
        number = "+1 " + matchUSAno1.group(1) + " " + matchUSAno1.group(2) + "-" + matchUSAno1.group(3)
    elif len(number) > 0 and number[0] != '+':
        regexParens = '[\(\)]'
        number = re.sub(regexParens, "", number)
        regexTopCountries = '^(30|31|33|41|44|49|55|60|61|65|82|86|90|380|852|886|966|971) ?(.*)'
        matchTopCountries = re.search(regexTopCountries, number)
        if matchTopCountries:
            number = matchTopCountries.group(1) + " " + matchTopCountries.group(2)
        # remove all leading 0's again
        regex0 = '^0+(.*)'
        match0 = re.search(regex0, number)
        regexDash = '-'
        number = re.sub(regexDash, " ", number)
        regexSpaces = ' {2,}'
        number = re.sub(regexSpaces, " ", number)
        number = "+" + number


    # POSTPROCESSING
    # remove the 0 from 440
    regex440 = '\+440'
    number = re.sub(regex440, "+44", number)
    regex0xx = '(\+[2-9][^ ].*\()0(.*)'
    match0xx = re.search(regex0xx, number)
    if match0xx:
        number = match0xx.group(1) + match0xx.group(2)
    # remove spaces from last 4 non-space characters
    number = remove_end_space(number, 4)
    # print out the original number and its formatted version
    if printing:
        if original != "None":
            print(original + " -> " + number)

    return number
#}}}

# Format all numbers in a column
# format_all_numbers(fileName, sheetName, col)
# Call using the command line
# format_all_numbess(filename, startRow, *cols
#{{{
# def format_all_numbers(fileName, sheetName, *cols):
def format_all_numbers(*args):
    args = args[0]
    fileName = args[1]
    startRow = int(args[2])
    cols = args[3:]
    if printing:
        print("Opening...")
    wb = openpyxl.load_workbook(args[1])
    # sheet = wb[sheetName]
    sheet = wb.worksheets[0]

    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')

    first = startRow
    last = sheet.max_row

    for col in cols:
        for row in range (first, last + 1):
            number = str(sheet[str(col) + str(row)].value)
            formatted = formatting_phone_number(number)
            if saving:
                sheet[col + str(row)].value = formatted

    if printing:
        print("Processing " + str((last + 1) - first) + " rows...")

    index = fileName[::-1].find('/')
    if printing and saving:
        fileName = fileName[:-index] + "formatted" + fileName[-index:]
        print("Saving " + fileName)
        wb.save(fileName)

    print(args)

    pygame.mixer.music.play()
    time.sleep(3)
    pygame.mixer.music.stop()
#}}}

# test stuff
#{{{
num1 = '+33 (0) 9490 3432'
num2 = '+33(0) 9490 3432'
num3 = '+33 (0)9490 3432'
num4 = '+33(0)9490 3432'

num5 = '+1 (202) 345-6789'
num6 = '(+1202) 345-6789'
num7 = '(+)1 202 3456789'
num8 = '+1 (202-345-6789'
num9 = '1202.345.6789'
num10 = '12023456789'
num11 = '2023456789'
num12 = '+1 (202) 345-6789 x2333'
num13 = '(202) 3456789'
num14 = '1 214-713 8014'
num15 = '(202) 345-6789'

num16 = '00 1(202)3456789'
num17 = '077 (202)3456789'
num18 = '078 (202)      3456789'

num19 = '(+1) 9173316874'
num20 = '+65 97970010 (Friday/weekend)'
num21 = '91 98455 39488'
num22 = 'Michael Austin (413) 668-6843'

# remove (0) from numbers
'''
print(formatting_phone_number(num1))
print(formatting_phone_number(num2))
print(formatting_phone_number(num3))
print(formatting_phone_number(num4))
'''
# format american numbers (must start with 1 or (ddd))
'''
print(formatting_phone_number(num5))
print(formatting_phone_number(num6))
print(formatting_phone_number(num7))
print(formatting_phone_number(num8))
print(formatting_phone_number(num9))
print(formatting_phone_number(num10))
print(formatting_phone_number(num11))
print(formatting_phone_number(num12))
print(formatting_phone_number(num13))
print(formatting_phone_number(num14))
print(formatting_phone_number(num15))
'''
# dealing with header 0's
'''
print(formatting_phone_number(num16))
print(formatting_phone_number(num17))
print(formatting_phone_number(num18))
'''
# misc
'''
print(formatting_phone_number(num19))
print(formatting_phone_number(num20))
print(formatting_phone_number(num21))
print(formatting_phone_number(num22))
'''
#}}}

format_all_numbers(sys.argv)
