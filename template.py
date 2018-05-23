import openpyxl
import pygame
import time

def template():
    # Create a new excel file
    out = openpyxl.Workbook()
    # Open the worksheet we want to edit
    outsheet = out.create_sheet("transactions")

    # if 'sheet' appears randomly we can delete it
    rm = out.get_sheet_by_name('Sheet')
    out.remove_sheet(rm)

    #################
    # DO STUFF HERE #
    #################
    outsheet['A1'].value = "DATA GOES HERE"

    # Save the file
    out.save("newFile.xlsx")






    # Open an existing excel file
    wb = openpyxl.load_workbook(fileName + ".xlsx")
    sheet = wb[sheetName]

    #################
    # DO STUFF HERE #
    #################
    sheet['A1'].value = "DATA GOES HERE"

    wb.save("betterFile.xlsx")

    # LMK when the script is done
    pygame.init()
    pygame.mixer.music.load('../note.mp3')
    pygame.mixer.music.play()
    time.sleep(3)
    pygame.mixer.music.stop()
