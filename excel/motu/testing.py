from datetime import datetime

line1 = '1/1/2009'
line2 = '1/1/2019'
dateObj1 = datetime.strptime(line1,'%m/%d/%Y')
dateObj2 = datetime.strptime(line2,'%m/%d/%Y')

print(dateObj1, dateObj2)

print(line1 < line2)
