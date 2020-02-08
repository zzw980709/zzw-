"""
# 读excel文件脚本
import xlrd              #导入模块
data = xlrd.open_workbook('电影.xlsx')    #打开电影.xlsx文件读取数据
table = data.sheets()[0]       #读取第一个（0）表单
#或者通过表单名称获取 table = data.sheet_by_name(u'Sheet1')
print(table.nrows)            #输出表格行数
print(table.ncols)            #输出表格列数
print(table.row_values(0))    #输出第一行
print(table.col_values(0))    #输出第一列
print(table.cell(0,2).value)  #输出元素（0,2）的值
"""
"""
#写excel文件脚本
import xlwt                            #导入模块
wb = xlwt.Workbook(encoding = 'ascii')  #创建新的Excel（新的workbook），建议还是用ascii编码
ws = wb.add_sheet('weng')               #创建新的表单weng
ws.write(0, 0, label = 'hello')         #在（0,0）加入hello
ws.write(0, 1, label = 'world')         #在（0,1）加入world
ws.write(1, 0, label = '你好')
wb.save('weng.xls')                     #保存为weng.xls文件
"""
"""
# 改exel文件
import xlrd                           #导入模块
from xlutils.copy import copy        #导入copy模块
rb = xlrd.open_workbook('weng.xls')    #打开weng.xls文件
wb = copy(rb)                          #利用xlutils.copy下的copy函数复制
ws = wb.get_sheet(0)                   #获取表单0
ws.write(0, 0, 'changed!')             #改变（0,0）的值
ws.write(8,0,label = '好的')           #增加（8,0）的值
wb.save('weng.xls')                    #保存文件
"""
