# coding=utf-8
# E:/python code/linked
from pymongo import MongoClient
import re
import xlwt

client = MongoClient("mongodb://127.0.0.1:27017/")
db = client.linked
collection = db['edu_job']


def export_data():
    # i = 0
    # edu = []
    ids = []
    job = []
    for item in collection.find():
        _id = item['_id']
        # edu_back = item['background_education']
        past_job = item['experience_past']
        current_job = item['experience_current']
        t = 2018
        # if len(edu_back) != 0:
        #     i += 1
        #     for temp in edu_back:
        #         pattern = re.compile(r'(.*?)(University|College|universidade)(.*?).*', flags=re.I)
        #         m = pattern.match(str(temp['school_name']))
        #         if temp['datetime'] != "NoneNone" and m is not None:
        #             years = get_year(temp['datetime'])
        #             years = int(years[0])
        #             if years < t:
        #                 if t != 2018:
        #                     edu.pop()
        #                 t = years
        #                 edu.append(years)
        #                 if _id not in ids:
        #                     ids.append(_id)
        # if len(past_job) != 0:
        #     for job1 in past_job:
        #         if job1['datetime'] != "NoneNone":
        #             year1 = get_year(job1['datetime'])
        #             job_years1 = int(year1[0])
        #             if job_years1 < t:
        #                 if t != 2018:
        #                     job.pop()
        #                 t = job_years1
        #                 job.append(t)
        #                 if _id not in ids:
        #                     ids.append(_id)
        # elif len(current_job) != 0:
        #     for job2 in current_job:
        #         if job2['datetime'] != "NoneNone":
        #             year2 = get_year(job2['datetime'])
        #             job_years2 = int(year2[0])
        #             if job_years2 < s:
        #                 if t != 2018:
        #                     job.pop()
        #                 s = job_years2
        #                 job.append(s)
        #                 if _id not in ids:
        #                     ids.appent(_id)
        if len(past_job) != 0:
            get_data(past_job, ids, job, _id, t)
        elif len(current_job) != 0:
            get_data(current_job, ids, job, _id, t)

    return job, ids

    # print("获得的记录数目：", i)
    # print(len(edu))
    # print(len(ids))
    # return edu, ids


def get_data(arr1, ids, data, _id, t):
    for job1 in arr1:
        if job1['datetime'] != "NoneNone":
            year1 = get_year(job1['datetime'])
            job_years1 = int(year1[0])
            if job_years1 < t:
                if t != 2018:
                    data.pop()
                t = job_years1
                data.append(t)
                if _id not in ids:
                    ids.append(_id)


def write_to_xls():
    # edu, ids = export_data()
    job, ids = export_data()
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    # print(type(book))
    sheet = book.add_sheet('age', cell_overwrite_ok=True)
    # print(type(sheet))
    sheet.write(0, 0, 'id')
    sheet.write(0, 1, 'F_job_year')
    # sheet.write(0, 2, 'age')
    # for i in range(len(edu)):
    #     age = 2018 - edu[i] + 18
    #     sheet.write(i + 1, 0, ids[i])
    #     sheet.write(i + 1, 1, edu[i])
    #     sheet.write(i + 1, 2, age)
    #     i += 1
    for i in range(len(ids)):
        sheet.write(i+1, 0, ids[i])
        sheet.write(i+1, 1, job[i])
        i += 1
    book.save(r'C:\Users\Administrator\Desktop\job.xls')


def write_to_mongodb():
    edu, ids = export_data()
    collection1 = db["id_age"]
    for i in range(len(edu)):
        age = 2018 - edu[i] + 18
        my_table = {"_id": ids[i], "edu_year": edu[i], "age": age}
        collection1.insert_one(my_table)


def get_year(datetime):
    str1 = str(datetime)
    # str1 = ''.join(str1.split(' '))
    # arr = str1.split("–")
    pattern = re.compile(r'([0-9][0-9][0-9][0-9])')
    arr = pattern.findall(str1)
    return arr


if __name__ == '__main__':
    write_to_xls()
