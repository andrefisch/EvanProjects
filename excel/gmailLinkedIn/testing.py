import re

number = "+44 (207) 333-2222"
regexParens = '[\(\)]'
number = re.sub(regexParens, "", number)
print(number)

