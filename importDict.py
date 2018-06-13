import sys

def importDict():
    # Open a file with sys.argv
    with open(sys.argv[1]) as f:
        dicty = {}
        for line in f:
            # split line by semicolon
            code, country = line.split(';')
            # do not take the \n at the end
            dicty[code] = country[:-1]
        return dicty

dicty = importDict()
