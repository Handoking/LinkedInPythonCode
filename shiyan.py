import numpy as np
import xlrd


def read_xls(path):
    xls_file = path  # 打开指定路径中的xls文件
    book = xlrd.open_workbook(xls_file)  # 得到Excel文件的book对象，实例化对象
    sheet1 = book.sheet_by_index(2)  # 通过sheet索引获得sheet对象
    col1_data = sheet1.col_values(1)
    # col1_data.pop(0)
    col1_data = [int(i) for i in col1_data]   # 将float型的列表数据转换成int
    col2_data = sheet1.col_values(2)
    # col2_data.pop(0)
    col2_data = [int(i) for i in col2_data]
    col_data = [a*b for a, b in zip(col1_data, col2_data)]
    col1_data2 = [c**2 for c in col1_data]
    sum_y = sum(col2_data)
    sum_x = sum(col1_data)
    sum_x2 = sum(col1_data2)
    sum_xy = sum(col_data)
    # print(sum_xy)
    return sum_x, sum_y, sum_xy, sum_x2, len(col1_data)
    # sheet_name1 = book.sheet_names()[0]  # 获得指定索引的sheet表名字
    # print(sheet_name1)
    # sheet1 = book.sheet_by_name(sheet_name1)
    # print(sheet2)
    # sheet1 = book.sheet_by_name(sheet_name)  # 通过sheet名字来获取，当然如果知道sheet名字就可以直接指定
    # nrows = sheet0.nrows  # 获取行总数
    # # 循环打印每一行的内容
    # for i in range(nrows):
    #     print
    #     sheet1.row_values(i)
    # row_data = sheet0.row_values(0)  # 获得第1行的数据列表
    # print
    # row_data
    # arr1 = np.array([5, 2, 3, 4])
    # arr1 = np.append(arr1, 8)
    # print(arr1)
    # arr2 = np.array([1, 2, 3, 4, 0])
    # a = sum(arr1*arr2)
    # print(sum(arr1))
    # print(a)


def line_fit():
    sum_x, sum_y, sum_xy, sum_x2, n = read_xls(r'C:\Users\handoking\Desktop\job_finish.xls')
    a = np.mat([[n, sum_x], [sum_x, sum_x2]])
    b = np.array([sum_y, sum_xy])
    return np.linalg.solve(a, b)


if __name__ == '__main__':
    a0, a1 = line_fit()
    print(a0, a1)
