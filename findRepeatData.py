import xlrd
from xlrd import xldate_as_tuple

def readFile():
    workBook = xlrd.open_workbook("");
    sheet = workBook.sheet_by_index(0);
    dateList = []
    j = 0
    for i in range(sheet.nrows):
        if i!= 0:
            sCell = sheet.cell(i, 0)
            date = xldate_as_tuple(sCell.value, 0)
            # dateList.append(date)
            dateList.append(date[3]*3600+date[4]*60+date[5])
    for i in range(len(dateList)):
        if i != 0 and dateList[i] - dateList[i-1]==1:
            j = j+1
            print(dateList[i] - dateList[i-1])
    print(j)
readFile()