# -*- coding: utf-8 -*-
#使用说明
#
#path是目标email路径(可以使用相对路径这样只需要把.py文件和目标email.xlsx文件放在一个文件夹,双击.py文件就可以得到想要的邮件)
#targetmail是你要提取出目标emal类型
#

import xlrd
import xlwt
import sys
reload(sys)

path = 'emails.xlsx'
targetmail = 'gmail.com'

sys.setdefaultencoding('utf8')
excel = xlrd.open_workbook(path)
sheet = excel.sheets()[0]
content=sheet.cell(3,0).value.encode('utf-8')
#print sheet.cell(4,0).value
#print content
#print sheet.nrows
file = xlwt.Workbook()
table = file.add_sheet(targetmail,cell_overwrite_ok=True)
tempnums= 0
for num in range(0,sheet.nrows):
    email=sheet.cell(num,0).value.encode('utf-8')
    if email.endswith(targetmail):
        print email
        table.write(tempnums,0,email.encode('utf-8'))
        tempnums += 1
file.save(targetmail+'.xls')