# -*- coding: utf-8 -*-

"""
Created on Mon May 27 14:48:06 2024

@author: jinlongshi
@email:  jinlongshi@stu.xju.edu.cn

"""

import urllib.request  # 发送请求
from urllib import parse  # URL编码
import json  # 解析json数据
from openpyxl import load_workbook  # 从Excel中读取镇街名称
from time import sleep
import openpyxl

nameList = []  # 创建一个列表用于接收数据
book = load_workbook("J:\Personal file\Software programmer\新大博达.xlsx")  # 打开文件
nameSheet = book["data"]  # 读取工作表
# 按行读取第一列，并存入列表：
for row in range(1, nameSheet.max_row + 1):
    nameList.append(str(nameSheet["A%d" % row].value))
    
print(nameList)

dict = {}
for i in nameList:
    # 拼接请求
    url1 = 'http://restapi.amap.com/v3/place/text?keywords=' + i + '&city=乌鲁木齐&output=json&offset=1&page=1&key=dcb33fc98e98d046a94071c2bc1456c4'
    # 将一些符号进行URL编码
    newUrl1 = parse.quote(url1, safe="/:=&?#+!$,;'@()*[]")
    # 发送请求
    response1 = urllib.request.urlopen(newUrl1)
    # 读取数据
    data1 = response1.read()
    # 解析json数据
    jsonData1 = json.loads(data1)
    # pois→0→location得到经纬度，写入字典
    dict[i] = jsonData1['pois'][0]['location']
    # 拆分字符串，逗号之前是经度，逗号之后是纬度
    locations = dict[i].split(",")

print(dict)

distanceList = []  # 创建一个列表用于接收数据
k = len(nameList)  # nameList列表中元素个数
# 遍历nameList列表
for m in range(k):
    subList = []  # 创建一个子列表用于接收每一条数据，主要是为了后面方便创建数组
    for n in range(k):
        # 从nameList中得到镇街的名称，作为键，获得dict中的经纬度
        origin = dict[nameList[m]]
        destination = dict[nameList[n]]
        # 拼接请求
        url2 = 'https://restapi.amap.com/v3/direction/driving?origin=' + origin + '&destination=' + destination + '&extensions=all&strategy=5&output=json&key=dcb33fc98e98d046a94071c2bc1456c4'
        print(url2)
        # 编码
        newUrl2 = parse.quote(url2, safe="/:=&?#+!$,;'@()*[]")
        # 发送请求
        response2 = urllib.request.urlopen(newUrl2)
        # 接收数据strategy
        data2 = response2.read()
        # 解析json文件
        jsonData2 = json.loads(data2)
        # 从json文件中提取距离
        distance = jsonData2['route']['paths'][0]['distance']
        # 将距离写入子列表
        distanceList.append(int(distance))
        # 看一下得到的数据，这一行有没有无所谓
        print(nameList[m], nameList[n], distance)
    sleep(0.2)

print(distanceList)

# 先打开我们的目标表格，再打开我们的目标表单
wb = openpyxl.load_workbook(r"J:\Personal file\Software programmer\各建筑距离.xlsx")
ws = wb['Sheet1']
# 取出distance_list列表中的每一个元素，openpyxl的行列号是从1开始取得，所以我这里i从1开始取
for i in range(1, len(distanceList) + 1):
    distance = distanceList[i-1]
    # 写入位置的行列号可以任意改变，这里我是从第2行开始按行依次插入第11列
    ws.cell(row=i, column=3).value = distance
#     print("操作成功")
#  保存操作
wb.save(r"J:\Personal file\Software programmer\各建筑距离.xlsx")