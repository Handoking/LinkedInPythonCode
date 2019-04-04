#!/usr/bin/python3

from pymongo import MongoClient
import re
import xlwt


client = MongoClient('localhost', 27017)
db = client.linked
collection = db['edu_job']
# datas = collection.find()
# for data in datas:
#     print(data['_id'])


def age_est():
    ids1 = []
    ids2 = []
    edu_year = []
    job_year = []
    i = 0
    j = 0
    for data in collection.find():
        t = 2019
        r = 2019
        id1 = data['_id']
        edu = data['background_education']
        past_job = data['experience_past']
        cur_job = data['experience_current']
        if len(edu) != 0:
            i += 1
            for temp in edu:
                pattern = re.compile(r'(.*?)(University|College|universidade)(.*?).*', flags=re.I)
                m = pattern.match(str(temp['school_name']))
                if temp['datetime'] != "NoneNone" and m is not None:
                    year = get_year(temp['datetime'])
                    years = int(year[0])
                    if years < t:
                        if t != 2019:
                            edu_year.pop()
                        t = years
                        edu_year.append(years)
                        if id1 not in ids1:
                            ids1.append(id1)

        # else:
        #     ids1.append(id1)
        #     edu_year.append(' ')
        if len(past_job) != 0:
            j += 1
            get_data(past_job, ids2, job_year, id1, r)
        elif len(cur_job) != 0:
            j += 1
            get_data(cur_job, ids2, job_year, id1, r)
        # else:
        #     ids2.append(id1)
        #     job_year.append(' ')
    return ids1, ids2, edu_year, job_year


def get_data(arr1, ids, data, id1, t):
    for job1 in arr1:
        if job1['datetime'] != "NoneNone":
            job = get_year(job1['datetime'])
            job_years1 = int(job[0])
            if job_years1 < t:
                if t != 2019:
                    data.pop()
                t = job_years1
                data.append(t)
                if id1 not in ids:
                    ids.append(id1)


def get_year(date_time):
    str1 = str(date_time)
    pattern = re.compile(r'([0-9][0-9][0-9][0-9])')
    arr1 = pattern.findall(str1)
    return arr1


def write_to_xls():
    ids1, ids2, edu_year, job_year = age_est()
    print(len(ids1), len(edu_year),len(ids2), len(job_year))
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('edu_job_year', cell_overwrite_ok=True)
    sheet.write(0, 0, 'id1')
    sheet.write(0, 1, 'edu_year')
    sheet.write(0, 2, 'id2')
    sheet.write(0, 3, 'job_year')
    # sheet.write(0, 2, 'age')
    # for i in range(len(edu)):
    #     age = 2018 - edu[i] + 18
    #     sheet.write(i + 1, 0, ids[i])
    #     sheet.write(i + 1, 1, edu[i])
    #     sheet.write(i + 1, 2, age)
    #     i += 1
    for i in range(len(ids1)):
        sheet.write(i+1, 0, ids1[i])
        sheet.write(i+1, 1, edu_year[i])
    for i in range(len(ids2)):
        sheet.write(i+1, 2, ids2[i])
        sheet.write(i+1, 3, job_year[i])
        i += 1
    book.save(r'C:\Users\Administrator\Desktop\edu_job_year.xls')


if __name__ == "__main__":
    write_to_xls()

