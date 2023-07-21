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
