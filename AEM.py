import xlrd
from xlutils.copy import copy


def model():
    age = []
    book = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\ejy_refer.xls')
    work_space = copy(book)
    sheet = book.sheet_by_index(0)
    data0 = sheet.col_values(1)
    # data0.pop(0)
    data1 = sheet.col_values(2)
    # data1.pop(0)
    while '' in data0:
        data0.remove('')
    while '' in data1:
        data1.remove('')
    data0 = [int(i) for i in data0]
    data1 = [int(i) for i in data1]
    for i in range(len(data0)):
        if data0[i] != 0:
            if data1[i] != 0:
                if data1[i] <= data0[i]:
                    age.append(aem2(data1[i]))
                else:
                    age.append(aem1(data0[i]))
            else:
                age.append(aem1(data0[i]))
        elif data1[i] != 0:
            age.append(aem2(data1[i]))
        else:
            return 0
    sheet3 = work_space.get_sheet(0)
    age = [str(i) for i in age]
    for i in range(len(age)):
        sheet3.write(i, 4, age[i])
    work_space.save(r'C:\Users\Administrator\Desktop\AEM.xls')


def aem2(x):
    return round(1875.53-0.9173*x-0.95+1)


def aem1(x):
    return round(2036.01-1.00006*x+0.13+1)


if __name__ == '__main__':
    model()



