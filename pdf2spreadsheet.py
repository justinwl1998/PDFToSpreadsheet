#   TODO: Run through every billing statement, taking only the regular
#   monthly charges, equipment and services, service fees, and taxes sections
#   
#   Then output them to a spreadsheet.

import PyPDF2
import os
import sys
import xlwt
from getRMCPage import getRMCPage
from getCharges import getCharges
from xlwt import Workbook

directory = 'billings'
currency_style = xlwt.XFStyle()
currency_style.num_format_str = "[$$-409]#,##0.00;-[$$-409]#,##0.00"

# main code starts here
print("Read individual file (0) or read all files (1)?\n")
choice = int(input())
if choice == 0:
    fileList = []
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        if not os.path.isdir(f):
            fileList.append(f)

    for i in range(len(fileList)):
        print(i, ": ", fileList[i])
    print("Select file to test: ")
    fileChoice = int(input())

    if fileChoice > len(fileList):
        print("Invalid selection")
        sys.exit()

    reader = PyPDF2.PdfReader(fileList[fileChoice])
    billPage = getRMCPage(reader.pages[0].extract_text().splitlines()) - 1
    pageAsText = reader.pages[billPage].extract_text().splitlines()

    print(getCharges(pageAsText))
    
elif choice == 1:
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
            print("reading: ", filename[-14:-4])
            reader = PyPDF2.PdfReader(f)
            billPage = getRMCPage(reader.pages[0].extract_text().splitlines()) - 1
            pageAsText = reader.pages[billPage].extract_text().splitlines()

            charges = getCharges(pageAsText)

            # write date in spreadsheet
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
            print()
                
    wb.save('output.xls')
else:
    print("Invalid choice")




    
    
