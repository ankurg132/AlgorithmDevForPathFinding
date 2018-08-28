import openpyxl

wb = openpyxl.load_workbook('Project Data.xlsx')

Data = []

for name in wb.sheetnames:
    sheet = wb[name]
    sheet_contents = []
    Cell = lambda r, c: sheet.cell(row = r, column = c).value
    records = Cell(1, 1)
    columns_per_record = Cell(1, 2)
    column = 1
    for record in range(records):
        row = 3
        column = column + columns_per_record
        current_record = []
        while(Cell(row, column - columns_per_record) != 'END'):
            Row = []
            for col in range(column - columns_per_record, column):
                value = Cell(row, col)
                Row.append(value)
            current_record.append(Row)
            row = row + 1
        sheet_contents.append(current_record)
    Data.append([name, sheet_contents])
'''
for sheet in Data:
    print(sheet[0])
    print()
    for record in sheet[1]:
        for row in record:
            print(row)
        print()
    print()
'''
try:
    print("Welcome to flight booking terminal\n1.Domestic flights\n2.International flights")
    choice = int(input("Please enter your choice: "))
    print()
    if choice not in [1, 2]:
        raise Exception("Choice out of range")
except:
    print("You have given invalid input")

if choice == 1:
    num = 0
    print("Welcome to Domestic flight section")
    for record in Data[1][1]:
        num = num + 1
        print('%s. %s'%(num, record[0][0]))
    try:
        countryChoice = int(input("Please choose your country: "))
        print()
        if countryChoice < 1 or countryChoice > num:
            raise Exception("Choice out of Range")
    except:
        print("You have given invalid input")
    print("Departure: ")
    for index in range(1, len(Data[1][1][countryChoice - 1])):
        print('%s. %s'%(index, Data[1][1][countryChoice - 1][index][0]))
    try:
        origin = int(input("Enter choice: "))
        print()
        if origin < 1 or origin > index:
            raise Exception("Choice out of range")
    except:
        print("You have given invalid input")
else:
    print("This section is under development")
