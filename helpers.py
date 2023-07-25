import PyPDF2
import os
import xlwt
from xlwt import Workbook

from random import randint

currency_style = xlwt.XFStyle()
currency_style.num_format_str = "[$$-409]#,##0.00;-[$$-409]#,##0.00"

# Gets the page where the regular monthly charges are listed,
# assuming the first page has an "at a glance" view

def getRMCPage(firstPage):
    # Get page number to find regular monthly charges
    for entry in firstPage:
        if "Regular monthly charges Page\xa0" in entry:
            modLen = len("Regular monthly charges Page\xa0")
            
            convertedEntry = entry[modLen+1:].split(' ')
            return int(convertedEntry[0])

# Once page where charges is found, read the charges and return
# a list of strings for each charge

def getCharges(pageText):
    chargeList = []
    found = False
    for i in range(len(pageText)):
        if pageText[i] == "Additional information":
            break
        if "Regular monthly charges" in pageText[i]:
            found = True
        if found:
            try:
                if '$' in pageText[i].split()[-1]:
                    chargeList.append(pageText[i])
            except IndexError:
                continue
    return chargeList

# Scans all PDF files and writes all charges to a spreadsheet, taking the
# title of the charge and how much was charged.
def outputToSpreadsheet(directory):
    if os.path.exists("output.xls"):
        os.remove("output.xls")
    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    colMod = 0 # for determining which column and row to write to on the spreadsheet
    rowMod = 4
    chargeDict = {}
    cellList = [] # Occupied cell list, for duplicate charge checking
    chargeIndex = 1
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        if os.path.isfile(f):
            cellList.clear()
            reader = PyPDF2.PdfReader(f)
            billPage = getRMCPage(reader.pages[0].extract_text().splitlines()) - 1
            pageAsText = reader.pages[billPage].extract_text().splitlines()

            charges = getCharges(pageAsText)

            # write date or file name in spreadsheet
            sheet1.write(rowMod, 0, filename[-14:-4])
            for i in range(len(charges)):
                tokens = charges[i].split(' ')
                chargeName = ' '.join(tokens[:-1])

                # charge must have an uppercase letter at beginning and must not end with a period
                # this may not work for other bills
                if chargeName[0].islower() or tokens[-1][-1] == '.':
                    continue
                
                if chargeName not in chargeDict:
                    chargeDict[chargeName] = chargeIndex
                    sheet1.write(3, chargeDict[chargeName], chargeName) # Write new charge on this reserved row
                    chargeIndex += 1

                try:
                    sheet1.write(rowMod, chargeDict[chargeName], float(tokens[-1][1:]), style=currency_style)
                    cellList.append(chargeDict[chargeName])
                except:
                    # if charge with duplicate name is found, add new column for it
                    alreadyFound = False
                    dupeIndex = 2
                    while chargeName + "(" + str(dupeIndex) + ")" in chargeDict:
                        # check to see if the row is already occupied by the searched duplicate charge column
                        if chargeDict[chargeName + "(" + str(dupeIndex) + ")"] not in cellList:
                            alreadyFound = True
                            break
                        dupeIndex += 1

                    if not alreadyFound:
                        chargeDict[chargeName + "(" + str(dupeIndex) + ")"] = chargeIndex
                        sheet1.write(3, chargeDict[chargeName + "(" + str(dupeIndex) + ")"], chargeName + "(" + str(dupeIndex) + ")")
                        chargeIndex += 1
                        
                    if tokens[-1][0] == '-':
                        sheet1.write(rowMod, chargeDict[chargeName + "(" + str(dupeIndex) + ")"], -1 * float(tokens[-1][2:]), style=currency_style)
                    else:    
                        sheet1.write(rowMod, chargeDict[chargeName + "(" + str(dupeIndex) + ")"], float(tokens[-1][1:]), style=currency_style)
                    cellList.append(chargeDict[chargeName + "(" + str(dupeIndex) + ")"])


            rowMod += 1
                
    wb.save('output.xls')
