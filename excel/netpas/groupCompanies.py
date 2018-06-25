import openpyxl
import pygame
import time

saving = True
printing = True
testing = True
low_threshold = 75
high_threshold = 100
min_word_len = 5
bad_words = ['korea', 'global', 'shipping', 'golden', 'petra', 'shanghai', 'qingdao', 'marita', 'univer', 'royal', 'prima', 'universal']

# Edit Distance Function
#{{{
# modified edit distance algorithm:
# - see if string one is a substring of the next string
#   - if it is there is a match
#   - it not look at edit distance for substring and substring of equal length
#     - if it is above a certain percentage or below a certain edit number we can safely assume they are a match. keep that string and try with next one

def edit_distance(word1, word2):
    if printing:
        print("Looking at '" + word1 + "' and '" + word2 + "'")
    word1 = word1.lower()
    word2 = word2.lower()
    len_1 = len(word1)
    len_2 = len(word2)

    # make sure shorter word is first
    if len_1 > len_2:
        word1, word2 = word2, word1
        len_1 = len_2

    if (len_1 >= min_word_len and word1 not in bad_words):
        if (word1 in word2):
            if printing:
                print ("Edit distance: 0")
                print ("Percent Match: 100")
            return (0, 100)
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
            return (edit_distance, percent_match)
    return (0, 0)


#}}}

if printing:
    print("Opening...")
wb = openpyxl.load_workbook("netpas.xlsx")
sheet = wb['Netpas']

pygame.init()
pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')

# if 'sheet' appears randomly we can delete it
# rm = wb.get_sheet_by_name('Sheet')
# wb.remove_sheet(rm)

first = 5
last = sheet.max_row
# first = 70
# last = 1005

raw_name = "C"
company_name = "D"
relevance = "E"


# all rows algorithm
# - the first row is automatically its own company
# - for all other rows take current company we are looking at and compare edit distance to it
#   - if distance is way off move store new company and keep going

curr = ""
count = 0
for row in range (first, last + 1):
    # for first row we take the name automatically
    if row == first:
        curr = sheet[raw_name + str(row)].value
        if testing:
            print("now in first row")
            print(sheet[raw_name + str(row)].value)
        # write the data
        if saving:
            sheet[company_name + str(row)] = curr
    else:
        print(row)
        edit, match = edit_distance(curr, sheet[raw_name + str(row)].value)
        if match > low_threshold and match <= high_threshold:
            sheet[company_name + str(row)] = curr
            sheet[relevance + str(row)] = match
            count = count + 1
        else:
            curr = sheet[raw_name + str(row)].value
            sheet[company_name + str(row)] = curr

if printing:
    print()
    print("Out of " + str(1 + last - first) + " companies there were " + str(count) + " duplicates")
    print("Saving...")

wb.save("netpas.xlsx")

if printing:
    print()
    print("Done!")

pygame.mixer.music.play()
time.sleep(5)
pygame.mixer.music.stop()


# edit_distance("zbc Shipping", "Abc Shipping")
# edit_distance("abc Shiping", "Abc Shipping Company, LTD")
# edit_distance("Abc Shipping Company, LTD", "abc Shiping")
# edit_distance("abc Shipping", "Abc Shipping Company, LTD")
