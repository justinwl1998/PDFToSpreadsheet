import PyPDF2
import os
import xlwt
from xlwt import Workbook

# Gets the page where the regular monthly charges are listed,
# assuming the first page has an "at a glance" view

currency_style = xlwt.XFStyle()
currency_style.num_format_str = "[$$-409]#,##0.00;-[$$-409]#,##0.00"

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
    colMod = 0
    rowMod = 4
    chargeDict = {}
    rowList = []
    chargeIndex = 1
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        if os.path.isfile(f):
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
                    sheet1.write(3, chargeDict[chargeName], chargeName)
                    chargeIndex += 1

                try:
                    sheet1.write(rowMod, chargeDict[chargeName], float(tokens[-1][1:]), style=currency_style)
                except:
                    # if charge with duplicate name is found, try to add new column,
                    # this is assuming there is at most only one other similar charge
                    # on the same billing
                    dupeIndex = 2
                    if chargeName + "(" + str(dupeIndex) + ")" not in chargeDict:
                        chargeDict[chargeName + "(" + str(dupeIndex) + ")"] = chargeIndex
                        sheet1.write(3, chargeDict[chargeName + "(" + str(dupeIndex) + ")"], chargeName + "(" + str(dupeIndex) + ")")
                        chargeIndex += 1
                        
                    if tokens[-1][0] == '-':
                        sheet1.write(rowMod, chargeDict[chargeName + "(" + str(dupeIndex) + ")"], -1 * float(tokens[-1][2:]), style=currency_style)
                    else:    
                        sheet1.write(rowMod, chargeDict[chargeName + "(" + str(dupeIndex) + ")"], float(tokens[-1][1:]), style=currency_style)


            rowMod += 1
                
    wb.save('output.xls')
