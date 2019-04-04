from pymongo import MongoClient
import xlwt
import xlrd

client = MongoClient('localhost', 27017)
db = client.linked
collection1 = db['edu_job_age']
collection2 = db['information']


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


def get_skills():
    book = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\AEM.xls')
    sheet = book.sheet_by_index(0)
    id1 = sheet.col_values(0)
    data = [int(i) for i in id1]
    data1 = [str(i) for i in data]
    skills_likes = []
    collection0 = db["skill_likes"]
    for _id in data1:
        # _id = int(_id)
        print(_id)
        dict1 = {}
        item = collection2.find_one({'_id': _id})
        temp = item['endorsement']
        skills = temp['skills_name_list']
        likes = temp['skills_num_list']  # 从文档中获取技能和点赞数
        if len(skills) != 0:
            my_table = {"member_id": _id}
            for i in range(len(skills)):  # 将技能和点赞数一一对应并放在字典中，然后存入数组
                if i >= len(likes):
                    dict1[skills[i]] = '0'
                    my_table[skills[i]] = '0'
                else:
                    dict1[skills[i]] = likes[i]
                    my_table[skills[i]] = likes[i]
            skills_likes.append(dict1)
            # print(my_table)
            collection0.insert(my_table)
    return skills_likes
    # print(len(skill_likes))
    # print(skill_likes[0])


if __name__ == "__main__":
    # read_write_mongo()
    skill_likes = get_skills()
