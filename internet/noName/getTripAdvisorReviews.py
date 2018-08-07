from bs4 import BeautifulSoup as bs
from urllib.request import Request, urlopen
import openpyxl
import os
# import pygame
import re
import sys
# import time

DATE   = "A"
STARS  = "B"
REVIEW = "C"
SOURCE = "D"

def scrapeTripAdvisor():
    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("transactions")

    # if 'sheet' appears randomly we can delete it
    rm = out['Sheet']
    out.remove(rm)

    # List the first page of the reviews (ends with "#REVIEWS") - separate the websites with ,
    url1 = "https://www.tripadvisor.com/Restaurant_Review-g60745-d321640-Reviews"
    url2 = "-No_Name_Restaurant-Boston_Massachusetts.html"
    # looping through each site until it hits a break
    request = Request(url1 + url2)
    json = urlopen(request).read()
    array = str(json)
    soup = str(bs(array))
    # print(soup)
    regexSpaces = ' +'
    soup = re.sub(regexSpaces, ' ', soup)
    # print(prettyHTML)

    #{{{
    ex = '''ss="expand_inline scrname" onclick="ta.trackEventOnPage('Reviews', 'clic      k', 'user_name_name_click')">Edward R</span></div><div class="location"><span class="expand_inl      ine userLocation">Newyork</span></div></div><div class="memberOverlayLink" data-anchorwidth="90      " id="UID_FB9381F99324C729A746EAD298171E69-SRC_585005890" onmouseover="widgetEvCall('handlers.i      nitMemberOverlay', event, this);"><div class="memberBadgingNoText"><span class="ui_icon pencil-      paper"></span><span class="badgetext">5</span></div></div></div></div></div><div class="ui_colu      mn is-9"><div class="innerBubble"><div class="wrap"><div class="rating reviewItemInline"><span       class="ui_bubble_rating bubble_50"></span><span class="ratingDate relativeDate" title="June 4,       2018">Reviewed today </span></div><div class="quote isNew"><a href="/ShowUserReviews-g60745-d32      1640-r585005890-No_Name_Restaurant-Boston_Massachusetts.html" id="rn585005890" onclick="ta.setE      vtCookie('Reviews','click','title',0,this.href);ta.util.cookie.setPIDCookie('0');"><span class=      "noQuotes">A must visit-</span></a></div><div class="prw_rup prw_reviews_text_summary_hsx" data      -prwidget-init="handlers" data-prwidget-name="reviews_text_summary_hsx"><div class="entry"><p c      lass="partial_entry">The place has great food but better yet a long history in that same locati      on. We spoke to a wonderful gentlemen Nick who has worked there for four decades. The story of       area, how the NO NAME came to be. So glad I visited and...<span class="taLnk ulBlueLinks" oncli      ck="widgetEvCall('handlers.clickExpand',event,this);">More</span></p></div></div><div class="pr      w_rup prw_reviews_vote_line_hsx" data-prwidget-deferred="defer'''
    # print(prettyHTML)
    #}}}
    regexReview = 'ui_bubble_rating bubble_(\d).*?relativeDate" title="(.*?)<.*?"partial_entry">(.*?)<'
    '''
    print(regexMatch[0][0])
    print(regexMatch[0][1])
    print(regexMatch[0][2])
    '''
    '''
    if regexMatch:
        print(regexMatch.group(1))
        print(regexMatch.group(2))
        print(regexMatch.group(3))
        '''

    outsheet[DATE   + '1'].value = "Date"
    outsheet[STARS  + '1'].value = "Rating"
    outsheet[REVIEW + '1'].value = "Review"
    outsheet[SOURCE + '1'].value = "Source"

    # probably want to start after 4607
    row = 2
    page = 0
    for page in range(0, 126):
        print("Now looking at page " + str(page))
        if page == 0:
            request = Request(url1 + url2)
        else:
            urlMiddle = "-or" + str(page * 10)
            request = Request(url1 + urlMiddle + url2)
        json = urlopen(request).read()
        array = str(json)
        soup = str(bs(array))
        regexMatch = re.findall(regexReview, soup)
        for i in range(0, len(regexMatch)):
            date = regexMatch[i][1]
            ind = date.find('>')
            date = date[ind + 10:]
            outsheet[DATE   + str(row)].value = date
            outsheet[STARS  + str(row)].value = int(regexMatch[i][0])
            outsheet[REVIEW + str(row)].value = regexMatch[i][2].replace('\\n', ' ').replace('amp;', '')
            outsheet[SOURCE + str(row)].value = "TripAdvisor"
            row = row + 1
        

    # Save the file
    out.save("tripAdvisorReviews.xlsx")

    '''
    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('/home/andrefisch/python/evan/note.mp3')
    pygame.mixer.music.play()
    time.sleep(5)
    pygame.mixer.music.stop()
    '''

scrapeTripAdvisor()
