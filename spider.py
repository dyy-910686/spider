# -*- codeing = utf-8 -*-
# Datatime:2020/12/5 5:04
# Filename:text3 .py
# Toolby: PyCharm
# @Author：邓育永


import time
import ssl
import xlwt
from selenium import webdriver
from bs4 import BeautifulSoup
ssl._create_default_https_context = ssl._create_unverified_context

#爬取网页，得到数据
def getData():
    chrome_driver = "src/chromedriver.exe"     #chromedriver驱动文件的位置

    browser = webdriver.ChromeOptions()
    browser.add_argument('user-agent=Mozilla / 5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, likeGecko) Chrome / 87.0.4280.66 Safari / 537.36')
    browser.add_argument('--ignore-certificate-errors')

    ss = webdriver.Chrome(executable_path=chrome_driver,chrome_options=browser)
    ss.get('https://data.stats.gov.cn/easyquery.htm?cn=C01')

    time.sleep(30)                                   #睡眠3秒，等待页面加载

    ss.find_element_by_id('mySelect_sj').click()    #点击时间的下拉列表框
    time.sleep(2)                                   #睡眠3秒，等待页面加载

    ss.find_element_by_class_name('dtText').send_keys('1949-,last10')   #在时间框里输入时间：1949-,last10
    time.sleep(1)                                                       #睡眠1秒，等待页面加载

    ss.find_element_by_class_name('dtTextBtn').click()  #点击确定
    time.sleep(1)                                       #睡眠1秒，等待页面加载


    ss.find_element_by_id('treeZhiBiao_4_a').click()    #点击人口
    time.sleep(3)                                       #睡眠3秒，等待页面加载

    ss.find_element_by_id('treeZhiBiao_30_a').click()   #点击总人口
    time.sleep(5)                                       #睡眠5秒，等待页面加载
    infos1 = ss.find_element_by_id('main-container')     #定位到id=main-container的元素（标签）
    thead = infos1.find_element_by_tag_name('thead').get_attribute("innerHTML")  #取得thead标签的html字符串，为str类型
    tbody1 = infos1.find_element_by_tag_name('tbody').get_attribute("innerHTML")  #取得tbody标签的html字符串，为str类型



    ss.find_element_by_id('treeZhiBiao_31_a').click()   #点击人口出生率、死亡率和自然增长率
    time.sleep(5)                                       #睡眠5秒，等待页面加载
    infos2 = ss.find_element_by_id('main-container')    #定位到id=main-container的元素（标签）
    tbody2 = infos2.find_element_by_tag_name('tbody').get_attribute("innerHTML")  # 取得tbody标签的html字符串，为str类型

    return thead, tbody1,tbody2

# 解析得到的数据
def analyseData():
    thead,tbody1,tbody2 = getData()                 #保存爬取到的网页源码,thead,tbody为str类型

    soup1 = BeautifulSoup(thead, "html.parser")     #将thead赋值给soup1，soup1为bs4.BeautifulSoup类型
    soup2 = BeautifulSoup(tbody1, "html.parser")    #将tbody赋值给soup2，soup2为bs4.BeautifulSoup类型
    soup3 = BeautifulSoup(tbody2, "html.parser")    #将tbody赋值给soup2，soup2为bs4.BeautifulSoup类型
    datalist = []                                   #声明一个存储所有数据的汇总列表

    for item in soup1.find_all(name='tr'):      #item为bs4.element.Tag类型
        bb = []                                 #临时存储年份
        for item2 in item.find_all(name='th'):  #item2为bs4.element.Tag类型
            bb.append(item2.text)   #item2.text为str类型
        bb.reverse()                #将所有数据重新排列，若不排列，则为2019 2018 2017... ，排序后为1949,1950,1951...
        for i in range(1):          #元素右移1位
            bb.insert(0, bb.pop())

        datalist.append(bb)         #将年份数据存入汇总列表中


    for item in soup2.find_all(name='tr'):      #item为bs4.element.Tag类型
        bb = []                                 #用于循环时，依次临时存储总人口、男性人口、女性人口、城镇人口、乡村人口
        for item2 in item.find_all(name='td'):  #item2为bs4.element.Tag类型
            bb.append(item2.text)               #item2.text为str类型
        bb.reverse()                            #将所有数据重新排列
        for i in range(1):                      #元素右移1位
            bb.insert(0, bb.pop())

        datalist.append(bb)                     #通过循环将总人口、男性人口、女性人口、城镇人口、乡村人口依次存入汇总列表中

    for item in soup3.find_all(name='tr'):      #item为bs4.element.Tag类型
        bb = []                                 #用于循环时，依次临时存储人口出生率、人口死亡率、人口自然增长率
        for item2 in item.find_all(name='td'):  #item2为bs4.element.Tag类型
            bb.append(item2.text)               #item2.text为str类型
        bb.reverse()                            #将所有数据重新排列
        for i in range(1):                      #元素右移1位
            bb.insert(0, bb.pop())              #通过循环将人口出生率、人口死亡率、人口自然增长率依次存入汇总列表中

        datalist.append(bb)

    return datalist


# 存储数据
def saveData():
    datalist = analyseData()    #获取数据解析好的数据

    book = xlwt.Workbook(encoding="utf-8", style_compression=0)            #新建工作本
    sheet = book.add_sheet("1949-2019年人口数据",cell_overwrite_ok=True)     #添加工作表

    for item in datalist:
        for item1 in item:
            print(item1)
            sheet.write(item.index(item1),datalist.index(item),item1)   #向工作表1949-2019年人口数据中写入数据

    sheet.write(0,0,"年份")       #将第一列的属性值改为‘年份’（爬虫得到是‘指标’
    book.save("source.xlsx")      #保存文件

    print("end")

#程序入口处
if __name__=='__main__':
    saveData()

