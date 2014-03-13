__author__ = 'Louis'
#-*- coding:utf-8 -*-
import lxml.html
from lxml.cssselect import CSSSelector
import urllib2
import os
import re
import xlsxwriter
from datetime import date
import glob
import codecs
import sys
import chardet


class DataExtractor:
    location = ''
    file_handler = None
    css_selector = None
    result = {}
    tree = None
    log_handler = None
    error_count = 0

    def __init__(self, location, css_selector, open_type):
        self.location = location
        url = ''
        if open_type == 'f':
            url = 'file:' + location
        self.file_handler = urllib2.urlopen(url)
        self.css_selector = css_selector
        self.tree = lxml.html.fromstring(self.file_handler.read().replace("<font>", '').replace("</font>", ''))

    def logger(self, msg):
        self.error_count += 1
        if self.error_count == 1:
            self.log_handler = codecs.open('Extract Error.txt', 'a+', 'gbk')
        self.log_handler.writelines(u"\r\n" + self.location + u"\r\n" + msg + u"\r\n")

    def get_result(self):
        for (n, s) in self.css_selector.items():
            result = CSSSelector(s[0])(self.tree)
            if n == u'ID':
                if len(result) == 0:
                    self.logger('CSS select fail')
            if s[1] == u'text':
                self.result[n] = ([x.text.replace(u'美元', '').replace(u'$', '') for x in result])
            elif s[1] == u'url':
                self.result[n] = ([x.get(u'href') for x in result])
            elif s[1] == u'img':
                self.result[n] = ([x.get(u'src').replace(u'%20', ' ') for x in result])
        return self.result


class FileLocator:
    root = []
    file_location = []
    log_handler = []
    error_count = 0

    def logger(self, msg):
        self.error_count += 1
        if self.error_count == 1:
            self.log_handler = open('File Locate Error.txt', 'a+')
        self.log_handler.write(msg + '\r\n')

    def __init__(self, root):
        encoding = sys.getfilesystemencoding()
        for dirpath, dirnames, filenames in os.walk(root):
            if len(dirnames) != 0:
                for filename in filenames:
                    if os.path.splitext(filename)[1] == '.htm' or os.path.splitext(filename) == '.html':
                        if len(glob.glob(os.path.join(dirpath, '*.txt'))) != 0:
                            temp = os.path.basename(glob.glob(os.path.join(dirpath, '*.txt'))[0])
                            self.file_location.append((dirpath.decode(encoding), filename.decode(encoding), temp.decode(encoding)))
                        else:
                            self.logger(dirpath + ':\r\n没有对应的contact文件' )

    def get_result(self):
        return self.file_location


class RegExtractor:
    location = ''
    file_handler = None
    reg = []
    log_handler = None
    error_count = 0
    content = ''
    result = {}

    def logger(self, msg):
        self.error_count += 1
        if self.error_count == 1:
            self.log_handler = codecs.open('Contact Info Error.txt', 'a+', 'gbk')
        self.log_handler.write('\r\n' + self.location + '\r\n' + msg + '\r\n')

    def __init__(self, location, reg):
        self.location = location
        self.reg = reg
        self.file_handler = open(location, 'r')
        self.content = self.file_handler.read()
        file_encoding = chardet.detect(self.content)
        file_encoding['encoding'] = file_encoding['encoding'] == 'GB2312' and 'gbk' or file_encoding['encoding']
        try:
            self.content = self.content.decode(file_encoding['encoding'])
        except Exception, e:
            self.logger(e.__str__())
    def get_result(self):
        warning = 0
        for n, r in self.reg.items():
            temp = re.findall(r, self.content)
            result = 'N/A'
            if len(temp) == 0:
                warning += 1
            else:
                result = temp[0]
            self.result[n] = result
        if warning > 0:
            self.logger(u'联系人格式错误')
        return self.result


class ExcelOutput:
    workbook = None
    worksheet = None
    error_count = 0
    text_format = None
    date_format = None
    money_format = None
    current_row = 0
    properties = {}
    files = None
    html_selector = {}
    contact_selector = {}
    html_data = {}
    contact_data = {}

    def __init__(self, properties, html_selector, contact_selector):
        self.properties = properties
        self.html_selector = html_selector
        self.contact_selector = contact_selector
        self.files = FileLocator('.').get_result()
        self.workbook = xlsxwriter.Workbook('sum.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.text_format = self.workbook.add_format()
        self.text_format.set_text_wrap()
        self.text_format.set_align('top')
        self.date_format = self.workbook.add_format({'num_format': 'yyyy/mm/dd'})
        self.date_format.set_align('top')
        self.money_format = self.workbook.add_format({'num_format': '$#,##0.00'})
        self.money_format.set_align('top')
        for n, prop in self.properties.items():
            self.worksheet.set_column(prop[1], prop[1], prop[0])
            self.worksheet.write(self.current_row, prop[1], n)
        self.current_row += 1

    def generate(self):
        for dirpath, html_file, contact_file in self.files:
            self.html_data = DataExtractor(os.path.join(dirpath, html_file), self.html_selector, 'f').get_result()
            self.contact_data = RegExtractor(os.path.join(dirpath, contact_file), self.contact_selector).get_result()
            for i in range(len(self.html_data.values()[0])):
                self.worksheet.set_row(self.current_row, 80)
                for n in self.html_data.keys():
                    if self.properties[n][2] == 'text':
                        self.worksheet.write_string(self.current_row, self.properties[n][1], self.html_data[n][i], self.text_format)
                    elif self.properties[n][2] == 'money':
                        self.worksheet.write_number(self.current_row, self.properties[n][1], float(self.html_data[n][i]), self.money_format)
                    elif self.properties[n][2] == 'date':
                        self.worksheet.write_number(self.current_row, self.properties[n][1], self.html_data[n][i], self.date_format)
                    elif self.properties[n][2] == 'img':
                        temp = os.path.join(dirpath, self.html_data[n][i])
                        if os.path.isfile(temp):
                            self.worksheet.insert_image(self.current_row, self.properties[n][1], temp)
                for n in self.contact_data.keys():
                    self.worksheet.write_string(self.current_row, self.properties[n][1], self.contact_data[n], self.text_format)
                self.worksheet.write_datetime(self.current_row, self.properties[u'订单日期'][1], date.fromtimestamp(os.path.getmtime(os.path.join(dirpath, html_file))), self.date_format)
                if i == len(self.html_data.values()[0]) -1:
                    self.worksheet.write_number(self.current_row, self.properties[u'总计'][1], sum(float(x) for x in self.html_data[u'价格']), self.money_format)
                self.current_row += 1

css_selector = {
    u'名称': ('.desc .name', 'text'),
    u'ID': ('.desc .sku', 'text'),
    u'链接': ('.item-desc .img a', 'url'),
    u'颜色': ('.primary-cart-content .item .color', 'text'),
    u'尺码': ('.primary-cart-content .item .size', 'text'),
    u'价格': ('.primary-cart-content .price .offer-price', 'text'),
    u'图片': ('.primary-cart-content .prod-img', 'img'),
}

reg_selector = {
    u'联系人': u"姓名[:：\s]*([a-zA-Z0-9_\u4e00-\u9fa5]+)\$*\n*",
    u'QQ': u"(?:QQ|qq)[:：\s]*([1-9][0-9]{4,})\$*\n*",
    u'旺旺': u"旺旺[:：\s]*([a-zA-Z0-9_\u4e00-\u9fa5]+)\$*\n*",
    u'电话': u"(?:手机|电话)[:：\s]*((?:\d{11})|(?:(?:\d{7,8})|(?:\d{4}|\d{3}|\d{5})-(?:\d{7,8})|(?:\d{4}|\d{3})-(?:\d{7,8})-(?:\d{4}|\d{3}|\d{2}|\d{1})|(?:\d{7,8})-(?:\d{4}|\d{3}|\d{2}|\d{1})))\$*\n*",
    u'地址': u"地址[:：\s]*(.+)\$*\n*",
}


properties = {
    u'名称': (15, 0, 'text'),
    u"ID": (10, 1, 'text'),
    u'链接': (15, 2, 'text'),
    u"颜色": (8, 3, 'text'),
    u"尺码": (8, 4, 'text'),
           #(u"库存", 8),
    u"价格": (8, 5, 'money'),
    u"联系人": (8, 6, 'text'),
    u"QQ": (12, 7, 'text'),
    u"旺旺": (15, 8, 'text'),
    u"电话": (15, 9, 'text'),
    u"地址": (20, 10, 'text'),
    u"订单日期": (10, 11, 'date'),
    u"总计": (8, 12, 'money'),
    u"图片": (14, 13, 'img')
}

if os.path.isfile('Extract Error.txt'):
    os.remove('Extract Error.txt')
if os.path.isfile('Contact Info Error.txt'):
    os.remove('Contact Info Error.txt')
if os.path.isfile('File Locate Error.txt'):
    os.remove('File Locate Error.txt')
ExcelOutput(properties, css_selector, reg_selector).generate()