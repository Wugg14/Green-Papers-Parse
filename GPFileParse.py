import openpyxl, pprint, datetime, csv

print('Opening workbook')

wb = openpyxl.load_workbook('G-politicians.xlsx')
sheet = wb['Incumbents and Candidates']

#To do: remove unnecessary rows, need only State abbreviation,
#office abbreviation and name
#remove candidates not up for election on 2018-11-06
#remove candidates not running for senate or house

newData = []

def writer(written):
    """ writes the new csv file """
    target= 'CandidateSummer2018.csv'
    myFile = open(target, 'w', newline='')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerows(written)

for row in range(2, sheet.max_row + 1):
    if sheet['D' + str(row)].value == 'S' or sheet['D' + str(row)].value == 'H':
        if sheet['G' + str(row)].value == datetime.datetime(2018, 11, 6, 0, 0):
            if sheet['K' + str(row)].value != 'Renominated' and sheet['K' + str(row)].value != 'on General Election ballot':
                state = sheet['C' + str(row)].value
                office = sheet['F' + str(row)].value
                first = sheet['M' + str(row)].value
                last = sheet['N' + str(row)].value
                party = sheet['Q' + str(row)].value
                newRow = [state, office, first, last, party]
                newData += [newRow]
    writer(newData)
    

