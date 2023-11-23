import modify
import getData
import shutil
import xlwt
import xlrd
import os
import collect
import match

dir = r'E:\study\零极点\2020.8.15数据采集2\material\data\泰州\泰州统计年鉴2010(EXCEL)308\www.shujuku.org'  # 某年鉴的表格文件夹
city = './泰州'  # 填入你在做的城市名称
year = '2011'  # 年鉴年份
baseexcel = r'E:\study\零极点\2020.8.15数据采集2\material\泰州\2012\match.xls'  # 调整之后的match.xls文件路径

if not os.path.exists(city):
    os.mkdir(city)
cityAndYear = city + "\\" + year
if not os.path.exists(cityAndYear):
    os.mkdir(cityAndYear)
excelDir = cityAndYear + "\\" + 'excel'  # 存放此年鉴所有表格的文件夹
collectDir = cityAndYear + "\\" + 'collect'  # 存放匹配到表名的表格的文件夹，即符合要求的表格，其中name.txt为存放各表格表名的文本文档，方便查看
modifyDir = cityAndYear + "\\" + 'modify'  # 存放调整格式后的表格的文件夹
dataExcel = cityAndYear + "\\" + 'data.xlsx'  # 存放爬取下来的数据的表格
matchExcel = cityAndYear + "\\" + 'match2012补.xls'  # 存放匹配到的数据

# collect.collectExcel(dir, excelDir, baseexcel, collectDir, pos=7, threshold=0.49, num=20)
# modify.modifyDir(collectDir, modifyDir)
# getData.getDataDir(modifyDir, dataExcel)

match.match2(baseexcel, dataExcel, matchExcel, year)  # 这一步运行非常耗时间，可考虑单独拿出来独立运行