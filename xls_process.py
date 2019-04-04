import xlrd
import xlwt


def xls_process(path):
    r_book = xlrd.open_workbook(path)
    w_book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = w_book.add_sheet('e_j_year', cell_overwrite_ok=True)
    sheet0 = r_book.sheet_by_index(0)
    # sheet1 = r_book.sheet_by_index(1)
    # sheet2 = r_book.sheet_by_index(2)
    col0 = sheet0.col_values(0)
    col1 = sheet0.col_values(1)
    col2 = sheet0.col_values(2)
    col3 = sheet0.col_values(3)
    col0.pop(0)
    col1.pop(0)
    col2.pop(0)
    col3.pop(0)
    while '' in col2:
        col2.remove('')
    while '' in col3:
        col3.remove('')
    while '' in col0:
            col0.remove('')
    while '' in col1:
        col1.remove('')
    col_0 = [int(i) for i in col0]
    col_1 = [int(i) for i in col1]
    col_2 = [int(i) for i in col2]
    col_3 = [int(i) for i in col3]
    print(type(col_0))
    print(type(col_2))
    # ids = []
    # _ids = []
    for k in col_2[:]:
        if k not in col_0:
            index = col_2.index(k)
            col_2.remove(k)
            col_3.pop(index)
    for j in col_0[:]:
        if j not in col_2:
            num = col_0.index(j)
            col_1.pop(num)
            col_0.remove(j)
    col_0 = [str(int(i)) for i in col_0]
    col_1 = [str(int(i)) for i in col_1]
    col_2 = [str(int(i)) for i in col_2]
    col_3 = [str(int(i)) for i in col_3]
    for i in range(len(col_0)):
        sheet.write(i, 0, col_0[i])
        sheet.write(i, 1, col_1[i])
        sheet.write(i, 2, col_3[i])
        sheet.write(i, 3, col_2[i])
    w_book.save(r'C:\Users\Administrator\Desktop\ejy_refer.xls')


if __name__ == '__main__':
    xls_process(r'C:\Users\Administrator\Desktop\edu_job_year.xls')