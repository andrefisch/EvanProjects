import re

# given an incomplete date string (Jun 2018) convert it to a date number -> (06/01/2018)
def fixDate(date):
    regexNum = '\d'
    if re.match(regexNum, date[0]):
        return date
    else:
        switcher = {
           "Jan": '01/01/' + date[4:],
           "Feb": '02/01/' + date[4:],
           "Mar": '03/01/' + date[4:],
           "Apr": '04/01/' + date[4:],
           "May": '05/01/' + date[4:],
           "Jun": '06/01/' + date[4:],
           "Jul": '07/01/' + date[4:],
           "Aug": '08/01/' + date[4:],
           "Sep": '09/01/' + date[4:],
           "Oct": '10/01/' + date[4:],
           "Nov": '11/01/' + date[4:],
           "Dec": '12/01/' + date[4:]
        }
        return switcher.get(date[:3], "Invalid month")

print(fixDate('Jan 2018'))
print(fixDate('01/17/2010'))
