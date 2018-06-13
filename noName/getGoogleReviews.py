from lxml import html
import openpyxl
import pygame
import re
import requests
import string
import time


def scrapeGoogleReviews():
    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("reviews")

    # if 'sheet' appears randomly we can delete it
    rm = out.get_sheet_by_name('Sheet')
    out.remove_sheet(rm)

    # google unique url endings for each restaurant
    url = 'https://www.google.com/search?ei=cC4QW4WfF4qN0wKKub6gBQ&q=no+name+restaurant+boston&oq=no+name+restaurant+boston&gs_l=psy-ab.3..35i39k1j0i67k1j0j0i20i263k1j0l2.5226.5832.0.5992.6.6.0.0.0.0.167.295.0j2.2.0.foo%2Ccfro%3D1%2Ckpnss%3D0...0...1.1.64.psy-ab..4.2.295....0.lzdByvtK2R0#lrd=0x89e37078246aed17:0xb9314acad62ae211,1,,,'
    num_reviews = 1500 # Number of reviews you want
    page_order = range(0, (num_reviews + 1), 20)
    count = 2
    pageNo = 0

    DATE   = "A"
    STARS  = "B"
    REVIEW = "C"


    outsheet[DATE   + '1'].value = "Date"
    outsheet[STARS  + '1'].value = "Rating"
    outsheet[REVIEW + '1'].value = "Review"

    # print("Currently on page number " + str(pageNo))
    page = requests.get(url)
    tree = html.fromstring(page.text)
    text = html.tostring(tree).decode("utf-8")
    print(text)

    index = text.find('<script type="application/ld+json">')
    text = text[index:]
    # find rating, date, and description for each review
    regexReview = 'reviewRating": \{"ratingValue":.*?(\d\.?\d?).*?datePublished": "(\d{4}-\d{2}-\d{2}).*?description": "(.*?)", "author'
    matchReview = re.findall(regexReview, text)

    for i in range (0, len(matchReview)):
        rating = int(matchReview[i][0])
        date   = matchReview[i][1]
        review = matchReview[i][2]
        print(rating, date, review)
        outsheet[DATE   + str(count)].value = date
        outsheet[STARS  + str(count)].value = rating
        outsheet[REVIEW + str(count)].value = review
        count = count + 1

    # works but all data comes in in a random order
    '''
    for i in range(0, len(dates)):
        print(dates[i].get("content"))
    for i in range(0, len(stars)):
        print(stars[i].get("content"))
    for i in range(0, len(reviews)):
        print(reviews[i])
    reviews = tree.xpath('//p[@itemprop="description"]/text()')
    dates = tree.xpath('//meta[@itemprop="datePublished"]')
    stars = tree.xpath('//meta[@itemprop="ratingValue"]')
    if reviews: # check if there is no review
        mod_reviews = []
        for rev in range(0, len(reviews)):
            count = count + 1
            if dates:
                outsheet[DATE   + str(count)].value = dates   [rev].get("content")
            if stars:
                outsheet[STARS  + str(count)].value = stars   [rev].get("content")
            # we already checked this one
            outsheet[REVIEW + str(count)].value = reviews [rev]
    '''

    # Save the file
    out.save("newFile.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

scrapeGoogleReviews()
