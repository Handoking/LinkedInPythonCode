import xlrd
from xlutils.copy import copy


def est_age(a, b):
    book = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\job_finish.xls')
    book2 = copy(book)
    # print(type(book2))
    sheet = book.sheet_by_index(2)
    data = sheet.col_values(1)
    # data.pop(0)
    while '' in data:
        data.remove('')
    data = [int(i) for i in data]
    age = []
    for x in data:
        age.append(a+b*x)
    sheet1 = book2.get_sheet(2)
    # print(sheet1)
    # sheet1.write(0, 3, 'est_age')
    for i in range(len(age)):
        sheet1.write(i, 8, age[i])
    book2.save(r'C:\Users\Administrator\Desktop\job_est.xls')


if __name__ == "__main__":
    est_age(2027.7138, -0.99254)
