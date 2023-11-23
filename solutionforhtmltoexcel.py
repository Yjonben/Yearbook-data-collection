import pandas as pd
from bs4 import BeautifulSoup


def transformone():
    # 这个程序处理杭州的htm很不错
    url = 'E:/寒假工作/城市能源数据采集/实验/上海HTML/C0609E.htm'
    tb = pd.read_html(url)[0]
    tb.to_csv(r'E:/寒假工作/城市能源数据采集/实验/实验运行结果数据/test1.csv', mode='a', encoding='utf_8_sig', header=1, index=0)


'''
感觉这个和上面的没什么区别
def test():
    url = ''
    df = pd.read_html(url, encoding='utf-8')
    bb = pd.ExcelWriter('out.xlsx')
    df[0].to_excel(bb)
    bb.close()
'''


def transformtwo(path, target):
    # 这个程序是我自己编的，可以处理杭州的HTML
    # url = 'E:/寒假工作/城市能源数据采集/实验/杭州HTML/showChapter.aspx-id=12-22.htm'
    url = path
    soup = BeautifulSoup(open(url, encoding='utf-8'), features='lxml')
    i1 = soup.find_all(id="Label1")
    ii1 = i1[0].string
    i2 = soup.find_all(id="Label2")
    ii2 = i2[0].string
    ii = ii1 + '' + ii2
    print(ii)
    tbl = pd.read_html(soup.prettify())[0]
    # print(tbl)
    lis=[]
    lis.append(ii)
    df1 = pd.DataFrame(lis)
    df2 = pd.DataFrame(tbl)
    data = pd.concat([df1, df2], axis=1, ignore_index=False)
    # data = data.T
    # data.to_excel('E:/寒假工作/城市能源数据采集/实验/实验运行结果数据/test2.xls', encoding='utf_8_sig')
    data.to_excel(target, encoding='utf_8_sig')


'''
def test():
    url = 'E:/寒假工作/城市能源数据采集/实验/杭州HTML/showChapter.aspx-id=12-22.htm'
    # html = requests.get(url)
    soup = BeautifulSoup(open(url, encoding='utf-8'), features='lxml')
    # content = soup.select("#myTable04")
    tbl = pd.read_html(soup.prettify())[0]
    print(tbl)
    data = pd.DataFrame(tbl)  # 这里不行，Must pass 2-d input
    data.to_excel('E:/寒假工作/城市能源数据采集/实验/实验运行结果数据/test2.xls', encoding='utf_8_sig')
    # print(content)
    # tbl = pd.read_html(html.prettify(), header=0)[0]  # prettify()优化代码,[0]从pd.read_html返回的list中提取出DataFrame
    # tbl.to_csv(r'E:/寒假工作/城市能源数据采集/实验/实验运行结果数据/test2.csv', mode='a', encoding='utf_8_sig', header=1, index=0)


# test()
'''


path = r'D:\project\energy data\南京\南京统计年鉴2014(HTML)3301\南京统计年鉴2014(HTML)\nongye\6-6.htm'
path = 'D:/project/energy data/杭州/杭州统计年鉴2016（网页版）/nj2016/showChapter.aspx-id=1-08.htm'
target = './test.xls'
transformtwo(path, target)
