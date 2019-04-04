#!/usr/bin/env python
# coding=utf-8


import re
import json


def txt2json():
    # 你的文件路径
    path = "port_result.txt"
    # 读取文件
    with open(path, 'r', encoding="utf-8") as file:
        # 定义一个用于切割字符串的正则
        seq = re.compile(",")
        result = []
        # 逐行读取
        for line in file:
            lst = seq.split(line.strip())
            item = {
                "ip": lst[0],
                "port": lst[1:]
            }
            result.append(item)
        print(type(result))
    # 关闭文件
    with open('txt2json.json', 'w') as dump_f:
        json.dump(result, dump_f)


if __name__ == '__main__':
    txt2json()
