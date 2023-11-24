import openpyxl as pyxl
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
import re

workBookPath = 'mediumAircraftData.xlsx'
newDataSheetPath = r'wildlifeAnalysis.xlsx'
print("Loading workbook...")

workBook = pyxl.load_workbook(workBookPath)
print("Workbook loaded, cleaning data...")

workSheet = workBook.active

wildlifeAnalysis = Workbook()

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

def writeToExcel(dataList, name, itemName, chartTitle) :
    '''
    Function that takes in a multi-dimensional list of data and sorts it alphabetically or ascending, (dependent on integer or string values.)
    The sorted list is then iterated through and placed in to a new excel sheet.
    Once the data placement is finished, the table of values is used to make a bar graph.
    :params dataList: multi-dimensional array with tuples consisting of a unique item and their occurrences within a dataset.
            name: name used for the new sheet being created for the data.
            itemName: name used for the unique items, header for column one, and title for the x-axis of the chart.
            chartTitle: specific name used for the title of the chart.
    '''
    wildlifeAnalysis.create_sheet(name)
    sheet = wildlifeAnalysis[name]
    sortedList = sorted(dataList)
    
    sheet.cell(1,1,itemName)
    sheet.cell(1,2,"INCIDENTS")
    
    row,column = 2,1
    for item in sortedList :
        sheet.cell(row,column,item[0])
        sheet.cell(row,column+1,item[1])
        row += 1
        
    quantity = Reference(sheet, min_col=2, max_col=2, min_row=1, max_row=(len(sortedList)+1))
    items = Reference(sheet, min_col=1, max_col=1, min_row=2, max_row=(len(sortedList)+1))
    
    chart = BarChart()
    chart.add_data(quantity, titles_from_data = True)
    chart.set_categories(items)
    chart.title = chartTitle
    chart.x_axis.title = itemName
    chart.y_axis.title = "COLLISIONS"
    sheet.add_chart(chart,"D1")
    
    return
    
def cleanData():
    animalsWorksheet, yearWorksheet, monthWorksheet, airlineWorksheet = workSheet["AF"][1:], workSheet["B"][1:], workSheet["C"][1:], workSheet["F"][1:]

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


    writeToExcel(topTenPercentAnimals, 'ChartForAnimals', 'ANIMALS', 'Animals Involved in Aircraft Collisions')
    writeToExcel(topAirlines, 'ChartForAirlines', 'AIRLINES', 'Airlines Involved in Aircraft Collisions')
    writeToExcel(allYears, 'ChartForYears', 'YEARS', 'Aircraft Collisions by Year')
    writeToExcel(allMonths, 'ChartForMonths', 'MONTHS', 'Aircraft Collisions by Month')

    del wildlifeAnalysis['Sheet']
    wildlifeAnalysis.save(newDataSheetPath)
    
cleanData()
