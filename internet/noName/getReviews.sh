# install dependencies
sudo pip3 install bs4 lxml openpyxl requests
sudo pip3 list --outdated --format=freeze | grep -v '^\-e' | cut -d = -f 1  | xargs -n1 sudo pip3 install -U
# get yelp reviews
echo "Getting Yelp reviews..."
python3 getYelpReviews.py
# get trip advisor reviews
echo "Getting Trip Advisor reviews..."
python3 getTripAdvisorReviews.py
# combine them
echo "Combining reviews..."
python3 combine.py yelpReviews.xlsx tripAdvisorReviews.xlsx
# rename everything
echo "Saving everything..."
DATE=`date +%Y%m%d`
mkdir "$DATE"
mv yelpReviews.xlsx $DATE/"$DATE"YelpReviews.xlsx
mv tripAdvisorReviews.xlsx $DATE/"$DATE"TripAdvisorReviews.xlsx
mv combined.xlsx $DATE/"$DATE"Combined.xlsx
echo "Done!"
