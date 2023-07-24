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
        if "Regular monthly charges" in pageText[i]:
            found = True
        if found:
            try:
                if '$' in pageText[i].split()[-1]:
                    chargeList.append(pageText[i])
            except IndexError:
                continue
    return chargeList
