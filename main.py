from copy import copy
from numpy import append
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import MergedCell
import json
import io
import os
def excel_to_jsons(excel_file):
    book = openpyxl.load_workbook(excel_file)
    sheet = book["Sheet1"]
    max_row = sheet.max_row
    max_column = sheet.max_column
    TitleList=[]
    LanList = []
    version_column = 0
    module_column = 0
    component_column = 0
    key_column = 0
    zh_column = 0 #--
    en_column = 0 #--
    fix_column = 0
    
    for i in range(1,max_column):
        temp = sheet.cell(1,i).value
        if temp!=None:
            TitleList.append({temp:i})
            if temp == 'version':version_column = i
            elif temp == 'module':module_column = i
            elif temp == '二级':component_column = i
            elif temp == '三级':key_column = i
            # elif temp == 'zh':zh_coloum = i
            # elif temp == 'en':en_coloum = i  
            elif temp == '修訂版':fix_column = i   
            elif i>max(1,key_column): LanList.append({temp:i})
    max_end_row = getMaxEndRow(sheet,module_column)
    print("生成{}种语言:".format(len(LanList)))#--
    for i in range(len(LanList)):
        lan_name = list(LanList[i].keys())[0]
        lan_column = list(LanList[i].values())[0]
        excel_to_json(sheet,lan_name,max_end_row,module_column,component_column,key_column,lan_column)
        print(lan_name.upper())#--
    book.close()
            
def getMaxEndRow(sheet,dependcolumn):
    max_row = sheet.max_row
    max_end_row = 0
    flag = ''
    for i in range(2,max_row):
        temp = sheet.cell(i,dependcolumn)
        if isinstance(temp,MergedCell)==False:
            # print(temp.value,flag)
            if temp.value != flag:
                flag = temp.value
                max_end_row = i   
    return max_end_row         


def excel_to_json(sheet,lan_name,max_end_row,module_column,component_coloum,key_column,lan_column):
    result = {}
    l1 = getSortList(sheet,2,max_end_row,module_column) 
    for index in range(len(l1)):
        subRes = []
        l2 = getSortList(sheet,l1[index]["start"],l1[index]["end"],component_coloum)
        for i in range(len(l2)):
            value = getLanDetail(sheet,l2[i]["start"],l2[i]["end"],key_column,lan_column)
            key = l2[i]['head']
            if key != None:
                subRes.append({key:value})
        result[l1[index]['head']] = copy(subRes)
    if not os.path.exists('output'):  # 是否存在这个文件夹 
        os.makedirs('output')  # 如果没有这个文件夹，那就创建一个
    json_file_name ='./output/'+lan_name + '_result.json'
    save_json_file(result, json_file_name)

def getSortList(sheet,start=2,stop=400,cloumn=3):
    tempList = []
    heads = []
    for row in range(start,stop):
        temp=sheet.cell(row,cloumn)
        if(isinstance(temp,MergedCell)==False):
            tempList.append(row)
        if(temp.value!= None):
            heads.append(temp.value)
    tempList = tempList[slice(0,len(heads)*2)]
    sortList = []
    tempItem = {}
    for index in range(len(tempList)):
        if index % 2 == 0:
            tempItem['head'] = heads[int(index/2)]
            tempItem['start'] = tempList[index]
        else:
            tempItem['end'] = tempList[index]
            sortList.append(copy(tempItem))
    return sortList

def getLanDetail(sheet,start,end,key_cloumn=5,lan_cloumn=7):
    result = {}
    for i in range(start,end):
        key = sheet.cell(i,key_cloumn).value
        if key != None:
            result[key] = transNull(sheet.cell(i,lan_cloumn).value)
    return result
def transNull(nul):
    if nul == None:
        return ''
    else:
        return nul
                  

def save_json_file(jd,json_file_name):
    file = io.open(json_file_name,'w',encoding='utf-8')
    txt = json.dumps(jd, indent=2, ensure_ascii=False)
    file.write(txt)
    file.close()
if '__main__'==__name__:
    # excel_to_jsons(u'testPro.xlsx','result2.json')
    excel_to_jsons(u'testPro.xlsx')