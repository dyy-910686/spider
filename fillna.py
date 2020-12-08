# -*- codeing = utf-8 -*-
# Datatime:2020/12/5 21:22
# Filename:fillna .py
# Toolby: PyCharm
# @Author：邓育永

import pandas

data = pandas.read_excel(r'source.xls',sheet_name=0)

#缺失值填充
data= data.interpolate()            #若前后都不为空时，用前后的均值填充，同时兼具向前填充的功能
data= data.fillna(method='bfill')   #向后填充

data.to_excel("target.xlsx")              #保存数据
print("保存成功")




