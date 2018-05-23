from PyPDF2 import PdfFileReader, PdfFileWriter
import openpyxl
import os
import pygame
import re
import sys
import time

# Importing multiple files
'''
with open(sys.argv[1]) as f:
    for line in f:
        string = line
'''

class Transaction(object):
    buyer   = ""
    seller  = ""
    address = ""
    price   = ""
    date    = ""
    
    def __init__(self, buyer, seller, address, price):
        self.buyer   = buyer
        self.seller  = seller
        self.address = address
        self.price   = price
        self.date    = date

def parsePage(pdf, num):
    with open(pdf, 'rb') as f:
        reader = PdfFileReader(f)
        contents = reader.getPage(num).extractText().split('\n')
        return contents

def decrypt_pdf(input_path, output_path, password):
    with open(input_path, 'rb') as input_file, \
        open(output_path, 'wb') as output_file:
        reader = PdfFileReader(input_file)
        reader.decrypt(password)

        writer = PdfFileWriter()

        for i in range(reader.getNumPages()):
            writer.addPage(reader.getPage(i))

        writer.write(output_file)

# Address before semicolon is buyer, after is seller. othewise just seller
# Returns buyer address then seller address
def splitAddressSemicolon(address):
    index = address.find(';')
    if index > -1:
        return (address[:index].strip(), address[index + 1:].strip())
    else:
        return ("", address)

def splitAddress(address):
    # get rid of multiple spaces in the string
    regexSpaces = ' +'
    address = re.sub(regexSpaces, " ", address)
    # Replace "c/o" with "Care of" and get trim string
    address = address.replace("c/o", 'Care of').strip()
    # Split the address into street, town, zipcode, lot info
    regexAddress = '(.*?), *([a-zA-Z/ ]*).*?(\d+) */? *(.*)'
    matchAddress = re.search(regexAddress, address)
    if matchAddress:
        return (matchAddress.group(1).strip(), matchAddress.group(2).strip(), matchAddress.group(3).strip(), matchAddress.group(4).strip())
    else:
        return ("", "", "", "")

'''
test1 = "c/o Ares Management  LLC 245 Park Ave. 42nd Fl.,   New York, N.Y. 10067; 105 W.  1st St., Boston 02127/Parcels  1/2, ID 0601173000"

one, two = splitAddressSemicolon(test1)
print(splitAddress(one))
print(splitAddress(two))
'''
'''
test1 = "c/o 70 Green St.,  Charlestown 02129"
test2 = "76/86 Harvard St./141  Harvard St., Chelsea/Everett  02150/Parcels I/II/III, ID  2200534000"
test3 = "580 Washington   St. #5B, Boston 02111/ Millennium Avery  Condominium "

print(splitAddress(test1))
print(splitAddress(test2))
print(splitAddress(test3))
'''

BUYER           = "A"
SELLER          = "B"
BUYER_ADDRESS   = "C"
BUYER_STREET    = "D"
BUYER_TOWN      = "E"
BUYER_ZIP_CODE  = "F"
SELLER_ADDRESS  = "G"
SELLER_STREET   = "H"
SELLER_TOWN     = "I"
SELLER_ZIP_CODE = "J"
SELLER_LOT      = "K"
PRICE           = "L"
DATE            = "M"

printing = True

def recordTransactions():
    # Open the file for editing
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("transactions")

    # if 'sheet' appears randomly we can delete it
    rm = out.get_sheet_by_name('Sheet')
    out.remove_sheet(rm)
    
    outsheet[BUYER           + '1'].value = "Buyer"
    outsheet[SELLER          + '1'].value = "Seller"
    outsheet[BUYER_ADDRESS   + '1'].value = "Buyer Address"
    outsheet[BUYER_STREET    + '1'].value = "Buyer Street"
    outsheet[BUYER_TOWN      + '1'].value = "Buyer Town"
    outsheet[BUYER_ZIP_CODE  + '1'].value = "Buyer Zip Code"
    outsheet[SELLER_ADDRESS  + '1'].value = "Seller Address"
    outsheet[SELLER_STREET   + '1'].value = "Seller Street"
    outsheet[SELLER_TOWN     + '1'].value = "Seller Town"
    outsheet[SELLER_ZIP_CODE + '1'].value = "Seller Zip Code"
    outsheet[SELLER_LOT      + '1'].value = "Lot"
    outsheet[PRICE           + '1'].value = "Price"
    outsheet[DATE            + '1'].value = "Date"

    # Keep track of the current row
    count = 2
    
    '''
    - find the word real estate
    - flatten array into a single string
    - find all instances of buyer, seller, address, price
    - put into spreadsheet
    '''

    if printing:
        print("Processing " + str(len(sys.argv) - 1) + " pdf files")

    for r in range(1, len(sys.argv)):
        if printing:
            print("Decrypting " + sys.argv[r] + "...")

        # Create a dTransactions file
        pages = 25
        decrypt_pdf(sys.argv[r], "dTransactions.pdf", "secret_password")
        with open("dTransactions.pdf", 'rb') as input_file:
            reader = PdfFileReader(input_file)
            pages = reader.getNumPages()

        # Turn all pages into an array
        contents = []
        for i in range (10, pages):
            contents = contents + parsePage('dTransactions.pdf', i)

        # Find the word real estate
        start = 0
        on = False
        for i in range (0, len(contents)):
            if "REAL ESTATE" in contents[i]:
                start = i
                on = True
                break

        # Flatten array
        listy = ' '.join(contents[i:])

        # Find each instance of buyer, seller, address, price
        regexTransactionInfo = 'Buyer ?:? ?(.*?)Seller ?:? ?(.*?)Address ?:? ?(.*?)Price ?:? ?([,$0-9]*)'
        matchTransactionInfo = re.findall(regexTransactionInfo, listy)
        if printing:
            print("Found " + str(len(matchTransactionInfo)) + " transactions")
        buyer   = ""
        seller  = ""
        address = ""
        price   = ""
        # store the info
        for i in range (0, len(matchTransactionInfo)):
            buyer   = matchTransactionInfo[i][0]
            seller  = matchTransactionInfo[i][1]
            address = matchTransactionInfo[i][2]
            price   = matchTransactionInfo[i][3][1:].replace(',', '')

            # Split address info
            buyer_address, seller_address = splitAddressSemicolon(address)
            buyer_info = splitAddress(buyer_address)
            seller_info = splitAddress(seller_address)

            regexNum = '(\d+)'
            matchNum = re.search(regexNum, sys.argv[r])
            if matchNum:
                date = matchNum.group(1)
                date = date[4:6] + "/" + date[6:] + "/" + date[:4]
            else:
                date = sys.argv[r]

            if re.match("^\d+$", price):
                price = int(price)

            outsheet[BUYER           + str(count)].value = buyer
            outsheet[SELLER          + str(count)].value = seller
            outsheet[BUYER_ADDRESS   + str(count)].value = buyer_address
            outsheet[BUYER_STREET    + str(count)].value = buyer_info[0]
            outsheet[BUYER_TOWN      + str(count)].value = buyer_info[1]
            outsheet[BUYER_ZIP_CODE  + str(count)].value = buyer_info[2]
            outsheet[SELLER_ADDRESS  + str(count)].value = seller_address
            outsheet[SELLER_STREET   + str(count)].value = seller_info[0]
            outsheet[SELLER_TOWN     + str(count)].value = seller_info[1]
            outsheet[SELLER_ZIP_CODE + str(count)].value = seller_info[2]
            outsheet[SELLER_LOT      + str(count)].value = seller_info[3]
            outsheet[PRICE           + str(count)].value = price
            outsheet[DATE            + str(count)].value = date

            # Increment the count
            count = count + 1

        # Delete the transactions PDF so we can create a new one
        os.remove("dTransactions.pdf")
    
    if printing:
        print("Saving...")

    # Save the file
    out.save("BJJ Transactions.xlsx")

    if printing:
        print("Done!")

    pygame.init()
    pygame.mixer.music.load('../note.mp3')
    pygame.mixer.music.play()
    time.sleep(3)
    pygame.mixer.music.stop()

recordTransactions()
