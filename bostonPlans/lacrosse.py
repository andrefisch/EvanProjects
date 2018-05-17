from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import openpyxl
import re
import pygame, time

xfile = openpyxl.load_workbook('uwlc.xlsx')
sheet = xfile.get_sheet_by_name('2021')
start = 2
end = 102
driver = webdriver.Chrome()
driver.get("https://www.uwlax.edu/info/directory/")
driver.switch_to.frame(0)

for row in range (start, end):
    lastName = sheet['A' + str(row)].value
    firstName = sheet['B' + str(row)].value
    inputElement = driver.find_element_by_id("search_criteria")
    inputElement.clear()
    inputElement.send_keys(firstName + " " + lastName)
    inputElement.send_keys(Keys.ENTER)
    html = driver.page_source
    p = re.compile('[\w\.]*@\w*uwlax\.edu')
    m = p.search(html)
    if '<table class="table table-striped color-h" id="resultsTable">' in str(html):
        try:
            sheet['C' + str(row)] = m.group()
            # Keep track of how close we are to being done
            print (str(format((row - start) / (end - start) * 100.00, '.2f')) + "%: " + m.group())
        except Exception:
            pass
xfile.save('test.xlsx')
