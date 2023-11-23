import modify
import getData
import shutil
import xlwt
import xlrd
import pandas as pd
import os
import collect
import match
import msoffcrypto


# xls改名
def rename():
    dir = r'E:\study\零极点\2020.8.15数据采集2\material\泰州\2017\excel'
    filelist = os.listdir(dir)
    for i in filelist:
        if i[-3:] == 'xls' or i[-4:] == 'xlsx':
            fail=xlrd.open_workbook(dir+'\\'+i)
            table=fail.sheets()[0]
            name=table.cell_value(0, 0)
            print(name)
            os.rename(dir+'\\'+i,dir+'\\'+name+'.xls')

# 破解excel密码
def decrypt():
    path = r'E:\study\零极点\2020.8.15数据采集2\material\data\泰州\泰州统计年鉴2019-EXCEL-186\www.shujuku.org\2020_3_16 21_58_50_N2020020035000163.xls'
    file = msoffcrypto.OfficeFile(open(path, 'rb'))  # 读取原文件
    file.load_key(password='VelvetSweatshop')  # 填入密码, 若能够直接打开, 则为默认密码'VelvetSweatshop'
    file.decrypt(open('decrypted.xls', 'wb'))  # 解密后保存为新文件

# rename()
# path=r'E:\study\零极点\2020.8.15数据采集2\material\泰州\2019\match.xls'
# table=pd.read_excel(path)
# print(str(table.iloc[841,7])=='nan')

excelDir = cityAndYear + "\\" + 'excel'  # 存放此年鉴所有表格的文件夹
collectDir = cityAndYear + "\\" + 'collect'  # 存放匹配到表名的表格的文件夹，即符合要求的表格，其中name.txt为存放各表格表名的文本文档，方便查看
path=r'E:\study\零极点\2020.8.15数据采集2\material\泰州\2014\collect\name.txt'
data = []
for line in open(path,"r",encoding='UTF-8'): #设置文件对象并读取每一行文件
    line=line[:-1]
    if line[-1] == ')' or line[-1] == '）':
        try:
            ind = line.index('（')
        except:
            try:
                ind = line.index('(')
            except:
                pass
            else:
                line = line[:ind]
        else:
            line = line[:ind]
    if line not in data:
        data.append(line)
