from lxml import html
import openpyxl
import pygame
import re
import requests
import string
import time


def scrapeYelpReviews()
    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("reviews")

    # if 'sheet' appears randomly we can delete it
    rm = out.get_sheet_by_name('Sheet')
    out.remove_sheet(rm)

    # Yelp unique url endings for each restaurant
    restaurants = ['no-name-restaurant-south-boston-2']
    start_urls = ['http://www.yelp.com/biz/%s' % s for s in restaurants]
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

    for ur in start_urls:
        for o in page_order:
            pageNo = pageNo + 1
            print("Currently on page number " + str(pageNo))
            page = requests.get(ur + ("?start=%s" % o))
            tree = html.fromstring(page.text)
            text = html.tostring(tree).decode("utf-8")

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

            break
                
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
