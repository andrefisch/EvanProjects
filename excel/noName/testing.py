# tripadvisor Scrapper - use this one to scrape hotels

# importing libraries
from bs4 import BeautifulSoup
import urllib
import os
import urllib.request

# creating CSV file to be used

file = open(os.path.expanduser(r"TripAdviserReviews.csv"), "wb")
file.write(
    b"Organization,Address,Reviewer,Review Title,Review,Review Count,Help Count,Attraction Count,Restaurant Count,Hotel Count,Location,Rating Date,Rating" + b"\n")

# List the first page of the reviews (ends with "#REVIEWS") - separate the websites with ,
url = "https://www.tripadvisor.com/Restaurant_Review-g60745-d321640-Reviews-No_Name_Restaurant-Boston_Massachusetts.html#REVIEWS"
# looping through each site until it hits a break
thepage = urllib.request.urlopen(url)
soup = BeautifulSoup(thepage, "html.parser")
for i in range(0, len(soup)):


file.close()
