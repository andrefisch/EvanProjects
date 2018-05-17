import re

number = '212-830-6660'

regexUSA    = '\D*(\+?1?)?\D*(\d{3})\D*(\d{3})\D*(\d{4})'
regexUSAno1 = '^(\(\d{3}\))\D*(\d{3})\D*(\d{4})'
matchUSA    = re.search(regexUSA,    number)
matchUSAno1 = re.search(regexUSAno1, number)

print(matchUSA, matchUSAno1)
