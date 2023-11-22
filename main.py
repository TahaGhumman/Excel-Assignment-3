import pandas as pd

df = pd.read_excel("smallAircraftData.xlsx", usecols = ['Species Name', 'Incident Year', 'Incident Month', 'Operator'])
dataFrame = df.dropna()
dataList = dataFrame.values.tolist()

years, months, airlines, animals = {}, {}, {}, {}

def countItems(dict, column) :
    for row in range(len(dataList)) :
        item = dataList[row][column]
        if item not in dict :
            dict[item] = 1
        else :
            dict[item] += 1

    return dict

def countAnimals(dict, column) :
    for row in range(len(dataList)) :
        fullName = dataList[row][column]
        commonName = fullName.split(' ')

        if commonName[-1] not in dict :
            dict[commonName[-1]] = 1
        else :
            dict[commonName[-1]] += 1

    return dict

def maxCount(dict) :
    for item in dict :
        if item == max(dict, key=dict.get) :
            return (item, dict[item])
    return

countItems(years, 0)
countItems(months, 1)
countItems(airlines, 2)
countAnimals(animals, 3)

print('')

print('Most Collisions: ', maxCount(years), maxCount(months), maxCount(airlines), maxCount(animals))

print('')