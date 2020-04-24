import time
import datetime

import xlrd
import xlwt
from xlrd import xldate_as_tuple
from xlutils.copy import copy

def read_excel():
    # 打开文件
    workBook = xlrd.open_workbook('baby 503 filled/503_3_3000_filled.xlsx');
    workBook_2 = xlrd.open_workbook('baby 503/503_3_3000 for test.xlsx');
    new_2 = copy(workBook_2)
    # 1.获取sheet的名字
    # 1.1 获取所有sheet的名字(list类型)
    # allSheetNames = workBook.sheet_names();
    # allSheetNames_2 = workBook_2.sheet_names();
    # print(allSheetNames);

    # 1.2 按索引号获取sheet的名字（string类型）
    # sheet1Name = workBook.sheet_names()[0];
    # sheet1Name_2 = workBook_2.sheet_names()[0];
    # print(sheet1Name);

    # 2. 获取sheet内容
    ## 2.1 法1：按索引号获取sheet内容
    sheet1_content1 = workBook.sheet_by_index(0); # sheet索引从0开始
    sheet1_content1_2 = workBook_2.sheet_by_index(0);  # sheet索引从0开始
    ## 2.2 法2：按sheet名字获取sheet内容
    # sheet1_content2 = workBook.sheet_by_name('Sheet1');

    # 3. sheet的名称，行数，列数
    print(sheet1_content1.name,sheet1_content1.nrows,sheet1_content1.ncols);
    print(sheet1_content1_2.name, sheet1_content1_2.nrows, sheet1_content1_2.ncols);

    # 4. 获取整行和整列的值（数组）
    rows = sheet1_content1.row_values(3); # 获取第四行内容
    # cols_1 = sheet1_content1.col_values(0).strftime('%H:%M:%S'); # 获取第三列内容
    cols_1_temp = sheet1_content1.col_values(0)
    # cols_1_temp.remove("date time")
    cols_1_date= []
    cols_2_date= []
    heartRateData = []
    respRateData = []
    cols_1_miss_date= []
    cols_2_miss_date = []
    cols_2 = sheet1_content1_2.col_values(0);  # 获取第三列内容
    for i in range(sheet1_content1.nrows):
        if i != 0:
            sCell = sheet1_content1.cell(i, 0)
            date = xldate_as_tuple(sCell.value, 0)
            cols_1_date.append(date)
            heartRateData.append(sheet1_content1.cell(i, 2).value)


    for i in range(sheet1_content1_2.nrows):
        if i != 0:
            sCell = sheet1_content1_2.cell(i, 0)
            date = xldate_as_tuple(sCell.value, 0)
            cols_2_date.append(date)
            respRateData.append(sheet1_content1_2.cell(i, 1).value)

    print("numerics miss:")
    for i in range(len(cols_1_date)):
        if cols_1_date[i] not in cols_2_date:
            if cols_1_date[i] not in cols_2_miss_date:
                cols_2_miss_date.append(i)
    print("total size is ",len(cols_2_miss_date))

    print("missing respiratory data:")
    for i in cols_2_date:
        if i not in cols_1_date:
            if i not in cols_1_miss_date:
                cols_1_miss_date.append(i)
                print(i)
    print("total size is ", len(cols_1_miss_date))


def changeIntoDate(tuple):
    stamp = datetime.datetime(2011, 1, 3, tuple[3], tuple[4], tuple[5])
    now = stamp + datetime.timedelta(seconds=1)
    return str(now)[-8:]


if __name__ == '__main__':
    # changeIntoDate((0,0,0,7,17,59))
    read_excel();