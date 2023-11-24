import openpyxl as pyxl
import re

workBookPath = 'smallAircraftData.xlsx'
print("Loading workbook...")

workBook = pyxl.load_workbook(workBookPath)
print("Workbook loaded, cleaning data...")

workSheet = workBook.active

def getSpeciesName(animalName: str):
    splitAnimalName = animalName.split(",")

    if len(splitAnimalName) == 1:
        splitAnimalName = re.split(r"[-;,.\s]\s*", splitAnimalName[0])

        
        if splitAnimalName[-1][-1] == "S":
            return splitAnimalName[-1][:-1]
        elif splitAnimalName[-1] == "GEESE":
            return "GOOSE"
        
        return splitAnimalName[-1]
    else:
        animals = []
        for animalNames in splitAnimalName:
            animalNames = re.split(r"[-;,.\s]\s*", animalNames)

            if animalNames[-1][-1] == "S":
                 animals.append(animalNames[-1][:-1])
            elif animalNames[-1] == "GEESE":
                animals.append("GOOSE")
            else:
                animals.append(animalNames[-1])
        return animals

def countAnimals(animalsWorksheet: dict):
    animalDict = {}
    for animalData in range(len(animalsWorksheet)): 
        fullName = animalsWorksheet[animalData].value

        if fullName == None or fullName == "": fullName = "UNKNOWN"
        
        commonName = getSpeciesName(fullName)

        if type(commonName) == list:
            for name in commonName:    
                if name not in animalDict:
                    animalDict[name] = 1
                else :
                    animalDict[name] += 1
        else:
            if commonName not in animalDict:
                animalDict[commonName] = 1
            else :
                animalDict[commonName] += 1

    return animalDict

def countItems(someWorksheet: dict):
    countedDict = {}
    for worksheetData in range(len(someWorksheet)):
        someValue = someWorksheet[worksheetData].value
        if someValue in countedDict:
            countedDict[someValue] += 1
        else:
            countedDict[someValue] = 1
    return countedDict

def returnHighestandTop(someDict: dict, threshold: int = 0):

    highestData = max(someDict, key=someDict.get)
    highestDataPoints = []
    topData = []
    percentage = threshold / 100

    for dataPoint in someDict:
        if someDict[dataPoint] == someDict[highestData]:
            highestDataPoints.append(dataPoint)
            topData.append((dataPoint, someDict[dataPoint]))
            
        elif someDict[dataPoint] > (someDict[highestData] * percentage):
            topData.append((dataPoint, someDict[dataPoint]))
    
    return highestDataPoints, topData
    
def cleanData():
    animalsWorksheet, yearWorksheet, monthWorksheet,airlineWorksheet  = workSheet["AF"][1:], workSheet["B"][1:], workSheet["C"][1:], workSheet["F"][1:]

    animalDict, yearDict, monthDict, airlineDict = countAnimals(animalsWorksheet), countItems(yearWorksheet), countItems(monthWorksheet), countItems(airlineWorksheet)

    highestAnimals, topTenPercentAnimals = returnHighestandTop(animalDict, 10)
    highestYear, allYears = returnHighestandTop(yearDict)
    highestMonth, allMonths = returnHighestandTop(monthDict)
    highestAirline, topAirlines = returnHighestandTop(airlineDict)

    if len(airlineDict) > 15:
        highestAirline, topAirlines = returnHighestandTop(airlineDict, 10)

    print(f"The animals(s) mostly involved in incidents are: {highestAnimals} with {animalDict[highestAnimals[0]]} incidents total.")
    print(f"The year(s) with most accidents are: {highestYear}, with {yearDict[highestYear[0]]} incidents total")
    print(f"The month(s) with most accidents are: {highestMonth}, with {monthDict[highestMonth[0]]} incidents total")
    print(f"The airline(s) mostly involved in accidents are: {highestAirline}, with {airlineDict[highestAirline[0]]} incidents total")
    
cleanData()