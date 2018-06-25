import re

# Extract first, last, and middle name from a name with more than 3 parts
# determine_names(listy)
def determine_names(name):
    listy   = []
    dicty   = {}
    lasty   = []
    middley = []

    # extract title
    regexAppellation = '^(Mr\.?|Mrs\.?|Ms\.?|Rev\.?|Hon\.?|Dr\.?|Capt\.?|Dcn\.?|Amb\.?|Lt\.?|MIDN\.?|Miss\.?|Fr\.?) (.*)'
    matchAppellation = re.search(regexAppellation, name, re.IGNORECASE)
    if matchAppellation:
        dicty['appellation'] = matchAppellation.group(1)
        name = matchAppellation.group(2)
    else:
        dicty['appellation'] = ""
    # extract suffix
    regexSuffix = '(.*) (Jr\.?|Sr\.?|Ph\.?D\.?|Esq\.?|II|III|USCG ?R?e?t?\.?|USMS)$'
    matchSuffix = re.search(regexSuffix, name, re.IGNORECASE)
    if matchSuffix:
        dicty['suffix'] = matchSuffix.group(2)
        name = matchSuffix.group(1)
    else:
        dicty['suffix'] = ""
    # split names
    regexName = "([\w+\.-]+)"
    while True:
        matchName = re.search(regexName, name)
        if matchName:
            listy.append(matchName.group(1))
            name = name.replace(matchName.group(1), "")
        else:
            break

    # first item is always first name at this point
    dicty['first_name']  = listy[0]
    dicty['middle_name'] = ""
    dicty['last_name']   = ""
    del listy[0]
    if len(listy) == 0:
        return dicty
    # - reverse list 
    # - take first name in reversed list (last name) and add it to last name list, delete it
    # - look at next name and see if it is capitalized
    #   - if not add to last name list, repeat
    #   - otherwise add this and rest to middle name list
    listy = listy[::-1]
    lasty.append(listy[0])
    del listy[0]
    lasts = True
    for i in range(0, len(listy)):
        if (not listy[i].istitle()) and lasts:
            lasty.insert(0, listy[i])
        else:
            lasts = False
            middley.insert(0, listy[i])

    dicty['middle_name'] = ' '.join(middley)
    dicty['last_name']   = ' '.join(lasty)
    return dicty
