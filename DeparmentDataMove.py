import openpyxl

class department():

    def __init__(self, Name, Apps_Received, Complete, Referred, Withdrawn, Admitted, Denied, Attempted):
        self.Name = Name
        self.Apps_Received = Apps_Received
        self.Complete = Complete
        self.Referred = Referred
        self.Withdrawn = Withdrawn
        self.Admitted = Admitted
        self.Denied = Denied
        self.Attempted = Attempted

    def prettyPrint(self):
        print "{} {} {} {} {} {} {} {}".format(self.Name,
                                         self.Apps_Received,
                                         self.Complete,
                                         self.Referred,
                                         self.Withdrawn,
                                         self.Admitted,
                                         self.Denied,
                                         self.Attempted)

    def getValues(self):
        return (self.Apps_Received, self.Complete, self.Referred,
                self.Withdrawn, self.Admitted, self.Denied, self.Attempted)


inputFile = open('programs.txt')
lines = inputFile.readlines()
inputFile.close()

programMap = {}

for line in lines:
    key = line.split(',')
    value = key[1].strip()
    key = key[0].strip()
    programMap[key] = value

for k in programMap:
    print "{}: {}".format(k, programMap[k])

wb = openpyxl.load_workbook(filename='firstFile.xlsx')
ws = wb['Sheet1']

departmentList = {}

for row in ws.rows:
    if row[0].value in programMap.keys():
        departmentList[programMap[row[0].value]] = (department(programMap[row[0].value],
                              row[1].value,
                              row[2].value,
                              row[3].value,
                              row[4].value,
                              row[5].value,
                              row[6].value,
                              row[7].value
                              ))

print len(departmentList)

for k in departmentList.keys():
    print k
    departmentList[k].prettyPrint()

wb = openpyxl.load_workbook(filename='secondFile.xlsx')
ws = wb.get_active_sheet()

for row in ws.rows:
    if row[0].value in departmentList.keys():
        dl = departmentList[row[0].value]
        row[1].value, row[2].value, row[3].value, row[4].value, row[5].value, row[6].value, row[7].value = dl.getValues()

wb.save(filename='secondFile.xlsx')
wb.close()
