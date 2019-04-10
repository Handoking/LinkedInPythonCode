import xlwt
from pymongo import MongoClient
import xlrd
import re
from xlutils.copy import copy


client = MongoClient('localhost', 27017)
db = client.linked
collection1 = db['edu_job_age']
collection2 = db['information']
collection4 = db['skills_likes_ages']


def read_write_mongo():
    book = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\AEM.xls')
    sheet = book.sheet_by_index(0)
    id1 = sheet.col_values(0)
    edu_year = sheet.col_values(1)
    job_year = sheet.col_values(2)
    aem_age = sheet.col_values(4)
    collection0 = db["edu_job_age"]
    for i in range(len(id1)):
        id1[i] = str(int(id1[i]))
        edu_year[i] = str(int(edu_year[i]))
        job_year[i] = str(int(job_year[i]))
        aem_age[i] = str(int(aem_age[i]))
        # print(type(id1[i]))
        # print(type(edu_year[i]))
        my_table = {"_id": id1[i], "edu_year": edu_year[i], "job_year": job_year[i], "age": aem_age[i]}
        collection0.insert_one(my_table)


def write_mongo(like_num):
    collection3 = db['skills_num']
    for i in range(len(like_num)):
        skill = re.sub(r'\.', '_', (like_num[i])[0])
        my_table = {'_id': str(i+1), skill: str((like_num[i])[1])}
        collection3.insert_one(my_table)


def get_skills():
    book = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\AEM.xls')
    sheet = book.sheet_by_index(0)
    id1 = sheet.col_values(0)
    data = [int(i) for i in id1]
    data1 = [str(i) for i in data]
    skills_likes = []
    collection0 = db["skills_likes"]  # 新建表，来存储id age skills likes等信息
    colt1 = db["edu_job_age"]
    for _id in data1:
        # _id = int(_id)
        print(_id)
        dict1 = {}
        item = collection2.find_one({'_id': _id})
        edu_job = colt1.find_one({'_id': _id})
        age = edu_job['age']
        # print(age)
        # dict1['_id'] = _id
        # dict1['age'] = int(age)
        temp = item['endorsement']
        skills = temp['skills_name_list']
        likes = temp['skills_num_list']  # 从文档中获取技能和点赞数
        for i in range(len(likes)):
            if likes[i] == '99+':
                likes[i] = '150'
        if len(skills) != 0:
            like = [int(i) for i in likes]
            # my_table = {"_id": _id, "age": age}
            for i in range(len(skills)):  # 将技能和点赞数一一对应并放在字典中，然后存入数组
                # skill = re.sub(r'\.', '_', skills[i])  # 字符串中包含点时不能作为key值，替换为下划线
                if i >= len(like):
                    dict1[skills[i]] = 0
                    # my_table[skill] = '0'
                else:
                    dict1[skills[i]] = like[i]
                    # my_table[skill] = likes[i]
            skills_likes.append(dict1)
            # print(my_table)
            # collection0.insert_one(my_table)
    return skills_likes
    # print(len(skill_likes))
    # print(skill_likes[0])


# def replace__():
#     string1 = "node.js.js.js"
#     skill = re.sub(r'\.', '_', string1)
#     print(skill)


def get_num(skills_likes):
    dict2 = {}
    dict3 = {}
    for i in range(len(skills_likes)):  # 循环遍历字典，技能表合并，点赞数累加
        key_arr = list((skills_likes[i]).keys())
        # print(type(key_arr))
        # print(type((skills_likes[i])[key_arr[0]]))
        if len(key_arr) != 0:
            if key_arr[0] in dict3:  # 获取顺序1的技能并计数
                dict3[key_arr[0]] = dict3[key_arr[0]] + 1
            else:
                dict3[key_arr[0]] = 0
            for j in range(len(key_arr)):
                if key_arr[j] in dict2:
                    dict2[key_arr[j]] = dict2[key_arr[j]] + (skills_likes[i])[key_arr[j]]
                else:
                    dict2[key_arr[j]] = (skills_likes[i])[key_arr[j]]
                # print(type(dict2[key_arr[j]]))
                # print(type((skills_likes[i])[key_arr[j]]))
                # print(type(dict2[key_arr[j]]))
    skill_likes_num = sorted(dict2.items(), key=lambda x: x[1], reverse=True)  #字典按值降序排序
    skill_num = sorted(dict3.items(), key=lambda x: x[1], reverse=True)
    # print(type(skill_likes_num))
    # for i in range(30):
    #     print(i, (skill_likes_num[i]))
    #     print(skill_num[i])
    return skill_num, skill_likes_num


# 获取某技能对应的id和年龄数据
# 获取点赞数较多的某技能的
def skill_age():
    age = {}
    collection0 = db['skills_num']
    data = collection0.find_one({'_id': '6'})  # 获取某一技能的名字
    skill = list(data.keys())  # 转换成列表便于查询
    data1 = collection4.find()
    for dict0 in data1:
        # print(type(i))
        keys = list(dict0.keys())
        if skill[1] == keys[2]:  # 找到某一特定顺序技能所在的位置，并将其用户的年龄存储到数组，并计算相同年龄的人数
            if dict0['age'] in age_keys:
                age[dict0['age']] += 1
            else:
                age[dict0['age']] = 1
        age_keys = list(age.keys())
    # print(age_keys)
    # 第一次新建文件：
        # book = xlwt.Workbook(encoding='utf-8', style_compression=0)  # 写入xls文档中
        # sheet = book.add_sheet('Litigation', cell_overwrite_ok=True)
        # sheet.write(0, 0, 'age')
        # sheet.write(0, 1, 'number')
    # 向xls文件中添加新的sheet表：
    book = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\age_num1.xls')
    work_space = copy(book)
    sheet = work_space.add_sheet('Microsoft Project', cell_overwrite_ok=True)
    sheet.write(0, 0, 'age')
    sheet.write(0, 1, 'number')
    for i in range(len(age_keys)):
        sheet.write(i+1, 0, age_keys[i])
        sheet.write(i+1, 1, age[age_keys[i]])
    work_space.save(r'C:\Users\Administrator\Desktop\age_num1.xls')


if __name__ == "__main__":
    # read_write_mongo()
    # skill_likes = get_skills()
    # replace__()
    # skills_num, skill_like_num = get_num(skill_likes)
    # write_mongo(skills_num)
    skill_age()
