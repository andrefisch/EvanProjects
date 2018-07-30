from lxml import html
import openpyxl
import pygame
import re
import requests
import string
import time


def scrapeYelpReviews():
    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("reviews")

    # if 'sheet' appears randomly we can delete it
    rm = out['Sheet']
    out.remove(rm)

    # Yelp unique url endings for each restaurant
    restaurants = ['no-name-restaurant-south-boston-2']
    start_urls = ['http://www.yelp.com/biz/%s' % s for s in restaurants]
    num_reviews = 1500 # Number of reviews you want
    page_order = range(0, (num_reviews + 1), 20)
    count = 2
    pageNo = 0

    DATE       = "A"
    STARS      = "B"
    REVIEW     = "C"
    SOURCE     = "D"
    NAME       = "E"
    LOCATION   = "F"
    FRIENDS    = "G"
    NUMREVIEWS = "H"


    outsheet[DATE       + '1'].value = "Date"
    outsheet[STARS      + '1'].value = "Rating"
    outsheet[REVIEW     + '1'].value = "Review"
    outsheet[NAME       + '1'].value = "Name"
    outsheet[LOCATION   + '1'].value = "Location"
    outsheet[FRIENDS    + '1'].value = "Friends"
    outsheet[NUMREVIEWS + '1'].value = "Number of Reviews"
    outsheet[SOURCE     + '1'].value = "Source"

    for ur in start_urls:
        for o in page_order:
            pageNo = pageNo + 1
            print("Currently on page number " + str(pageNo))
            page = requests.get(ur + ("?start=%s" % o))
            tree = html.fromstring(page.text)
            text = html.tostring(tree).decode("utf-8")

            index = text.find('<script type="application/ld+json">')
            # index = text.find('<a class="user-display-name')
            text = text[index:]
            # find rating, date, and description for each review
            regexReview = 'reviewRating": \{"ratingValue":.*?(\d\.?\d?).*?datePublished": "(\d{4}-\d{2}-\d{2}).*?description": "(.*?)", "author'
            # regexRater = 'id="dropdown.*?>(.*?)</a>.*?<b>(.*?)</b>.*?<b>(.*?)</b>.*?<b>(.*?)</b>'
            # regexRater = 'id="dropdown.*?>(.*?)</a>.*?<b>'
            # regexRater = 'a class="user-display-name.*?>(.*?)</a>'
            # regexFriends = '<b>(.*)</b>.*?" friends'

            matchReview  = re.findall(regexReview, text)
            # matchRater   = re.findall(regexRater, texty)
            # matchFriends = re.search(regexFriends, text)
            for i in range (0, len(matchReview)):
                '''
                if matchRater:
                    print(matchRater.group(1))
                if matchFriends:
                    print(matchFriends.group(1))
                    '''
                '''
                name     = matchRater[i][0]
                location = matchRater[i][1]
                friends  = matchRater[i][2]
                reviews  = matchRater[i][3]
                outsheet[NAME       + str(count)].value = name
                outsheet[LOCATION   + str(count)].value = location
                outsheet[FRIENDS    + str(count)].value = friends
                outsheet[NUMREVIEWS + str(count)].value = reviews
                print(name, location, friends, reviews)
                '''
                rating   = int(matchReview[i][0])
                date     = matchReview[i][1]
                review   = matchReview[i][2]
                print(rating, date, review)
                outsheet[DATE   + str(count)].value = date
                outsheet[STARS  + str(count)].value = rating
                outsheet[REVIEW + str(count)].value = review
                outsheet[SOURCE + str(count)].value = "Yelp"
                count = count + 1


    # Save the file
    out.save("yelpReviews.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()

scrapeYelpReviews()
