import modify
import getData
import shutil
import xlwt
import xlrd
import os
import collect
import match

dir = r'E:\study\零极点\2020.8.15数据采集2\material\data\泰州\泰州统计年鉴2019-EXCEL-186\www.shujuku.org\分表'  # 某年鉴的表格文件夹
city = './泰州'  # 填入你在做的城市名称
year = '2019'  # 年鉴年份
baseexcel = '统一目录.xlsx'  # 统一目录

if not os.path.exists(city):
    os.mkdir(city)
cityAndYear = city + "\\" + year
if not os.path.exists(cityAndYear):
    os.mkdir(cityAndYear)
excelDir = cityAndYear + "\\" + 'excel'  # 存放此年鉴所有表格的文件夹
collectDir = cityAndYear + "\\" + 'collect'  # 存放匹配到表名的表格的文件夹，即符合要求的表格，其中name.txt为存放各表格表名的文本文档，方便查看
modifyDir = cityAndYear + "\\" + 'modify'  # 存放调整格式后的表格的文件夹
dataExcel = cityAndYear + "\\" + 'data.xls'  # 存放爬取下来的数据的表格
matchExcel = cityAndYear + "\\" + '限额以上住宿和餐饮法人企业主要财务指标match.xls'  # 存放匹配到的数据

# collect.collectExcel(dir, excelDir, baseexcel, collectDir)
# modify.modifyDir(collectDir, modifyDir)
# getData.getDataDir(modifyDir, dataExcel)

# match.match(baseexcel, dataExcel, matchExcel, year)  # 这一步运行非常耗时间，可考虑单独拿出来独立运行