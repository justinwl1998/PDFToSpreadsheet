# Gets the page where the regular monthly charges are listed,
# assuming the first page has an "at a glance" view

def getRMCPage(firstPage):
    # Get page number to find regular monthly charges
    for entry in firstPage:
        if "Regular monthly charges Page\xa0" in entry:
            modLen = len("Regular monthly charges Page\xa0")
            
            convertedEntry = entry[modLen+1:].split(' ')
            return int(convertedEntry[0])
