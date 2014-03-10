__author__ = 'Louis'
# coding=gbk
import lxml.html
from lxml.cssselect import CSSSelector
import urllib2
import os
import re
import xlsxwriter
from datetime import date
import glob

def getInfo(info):
    if len(info) != 0:
        return info[0]
    else:
        return 'N/A'

files = []
Columns = ((u"名称", 15),
           (u"ID", 10),
           (u"链接", 15),
           (u"颜色", 8),
           (u"尺码", 8),
           #(u"库存", 8),
           (u"价格", 8),
           (u"联系人", 5),
           (u"QQ", 12),
           (u"旺旺", 15),
           (u"电话", 15),
           (u"地址", 20),
           (u"订单日期", 10),
           (u"总计", 8),
           (u"图片", 14))
current_row = 1
workbook = xlsxwriter.Workbook('sum.xlsx')
worksheet = workbook.add_worksheet()
money_format = workbook.add_format({'num_format': '$#,##0.00'})
money_format.set_align('top')
myformat = workbook.add_format()
myformat.set_align('top')
myformat.set_text_wrap()
dateformat = workbook.add_format({'num_format': 'yyyy/mm/dd'})
dateformat.set_align('top')
for index, column in enumerate(Columns):
    worksheet.set_column(index, index, column[1])
    worksheet.write(0, index, column[0])

for dirpath, dirnames, filenames in os.walk('.'):
    for filename in filenames:
        if os.path.splitext(filename)[1] == '.htm' or os.path.splitext(filename) == '.html':
            #if os.path.isfile(os.path.join(dirpath, 'Contact.txt')):
            if len(glob.glob(os.path.join(dirpath, '*.txt'))) != 0:
                files.append((dirpath, filename))
for f in files:
    u = urllib2.urlopen('file:' + os.path.join(f[0], f[1]))
    content = u.read()
    tree = lxml.html.fromstring(content)
    sel_Name = CSSSelector('.desc .name')
    sel_OrderID = CSSSelector('.desc .sku')
    sel_Link = CSSSelector('.item-desc .img a')
    sel_Color = CSSSelector('.primary-cart-content .item .color')
    sel_Size = CSSSelector('.primary-cart-content .item .size')
    #sel_Status = CSSSelector('.primary-cart-content .item .status')
    sel_Price = CSSSelector('.primary-cart-content .price .offer-price')
    sel_Img = CSSSelector('.primary-cart-content .prod-img')

    rName = sel_Name(tree)
    rOrderID = sel_OrderID(tree)
    rLink = sel_Link(tree)
    rColor = sel_Color(tree)
    rSize = sel_Size(tree)
    #rStatus = sel_Status(tree)
    rPrice = sel_Price(tree)
    rImg = sel_Img(tree)
    Name = [x.text for x in rName]
    OrderID = [x.text for x in rOrderID]
    Link = [x.get('href') for x in rLink]
    Color = [x.text for x in rColor]
    Size = [x.text for x in rSize]
    print Name
    print OrderID
    print Link
    print Color
    print Size
    print os.path.join(f[0], f[1]).decode('gbk')
    #Status = ["".join("".join(x.text.split('\t')).split('\n')) for x in rStatus]
    Price = [float(x.text.replace('$', '')) for x in rPrice]
    Img = [x.get('src') for x in rImg]
    Total = sum(Price)
    print Name
    print OrderID
    print Link
    print Color
    print Size
    #print Status
    print Price
    print Img
    contact_file = glob.glob(os.path.join(f[0], '*.txt'))
    if len(contact_file) != 0:
        aa = contact_file[0].split('\\')
        aa = aa[len(aa) - 1]
        fileHandle = open(os.path.join(f[0], aa))
        info = fileHandle.read().decode('gbk')
        print info
        Contact_Name = getInfo(re.findall(u"姓名[:：](.+?)\n", info))
        Contact_QQ = getInfo(re.findall(u"QQ[:：](.+?)\n", info))
        Contact_WW = getInfo(re.findall(u"旺旺[:：](.+?)\n", info))
        Contact_Phone = getInfo(re.findall(u"手机[:：](.+?)\n", info))
        Contact_Address = getInfo(re.findall(u"地址[:：](.+)\n?", info))
        Order_Time = date.fromtimestamp(os.path.getctime(os.path.join(f[0], f[1])))
        print Order_Time
        print Contact_Name
        print Contact_QQ
        print Contact_WW
        print Contact_Phone
        print Contact_Address
    index = 0
    for N, O, L, C, S, P, I in zip(Name, OrderID, Link, Color, Size, Price, Img):
        worksheet.set_row(current_row, 80)
        worksheet.write_string(current_row, 0, N, myformat)
        worksheet.write_string(current_row, 1, O, myformat)
        worksheet.write_string(current_row, 2, L, myformat)
        worksheet.write_string(current_row, 3, C, myformat)
        worksheet.write_string(current_row, 4, S, myformat)
    #    worksheet.write_string(current_row, 5, Sta, myformat)
        worksheet.write_number(current_row, 6, P, money_format)
        worksheet.write_string(current_row, 7, Contact_Name, myformat)
        worksheet.write_string(current_row, 8, Contact_QQ, myformat)
        worksheet.write_string(current_row, 9, Contact_WW, myformat)
        worksheet.write_string(current_row, 10, Contact_Phone, myformat)
        worksheet.write_string(current_row, 11, Contact_Address, myformat)
        worksheet.write_datetime(current_row, 12, Order_Time, dateformat)
        if index == len(zip(Name, OrderID, Link, Color, Size, Price, Img)) - 1:
            worksheet.write_number(current_row, 13, Total, money_format)
        worksheet.insert_image(current_row, 14, os.path.join(f[0], I.replace('%20', ' ')))
        current_row += 1
        index += 1
workbook.close()