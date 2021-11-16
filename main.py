#!/usr/bin/python
# -*- coding: utf-8 -*-
import re
import time
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import io
import sys


def login(gid):
    driver = webdriver.Chrome(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')
    driver.get('https://qun.qq.com/member.html#gid=%s' % gid)
    driver.maximize_window()
    time.sleep(3)
    for i in range(1000):
        js = "window.scrollTo(0,document.body.scrollHeight)"
        driver.execute_script(js)
    res = driver.page_source
    driver.close()
    return res


def dispose(res):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='gb18030')
    soup = BeautifulSoup(res, 'lxml')
    c = soup.find_all('tr', attrs={"class": re.compile('mb')})
    list_a = []
    for i in c:
        str_a = i.text.replace('\n', '').replace('\t', ',')
        data = re.split(r',', str_a)
        data_list = [i for i in data if i != '']
        img = "https:" + i.img.get('src')
        data_list.insert(2, img)
        del data_list[0]
        if len(data_list) < 9:
            data_list.insert(2, '')
        list_a.append(data_list)
    return list_a


def write_execl(list_a):
    if len(list_a) > 2:

        # 创建execl
        new_time = time.strftime("%Y-%m-%d %H_%M_%S", time.localtime())
        workbook = xlsxwriter.Workbook('{}.xlsx'.format(new_time))
        worksheet = workbook.add_worksheet('sheet1')

        bold = workbook.add_format({
                            'bold': 1,  # 字体加粗
                            'fg_color': 'green',  # 单元格背景颜色
                            'align': 'center',  # 对齐方式
                            'valign': 'vcenter',  # 字体对齐方式
                            })
        work_header = ['QQ昵称', '头像地址', '群昵称', 'QQ号', '性别', 'Q龄', '入群时间', '等级（积分）', '最后发言']
        worksheet.write_row('A1', work_header, bold)
        for index in range(len(list_a)):
            worksheet.write_row('A%s' % (2 + index), list_a[index])
        workbook.close()
    else:
        print('请检查群号是否有误，没有获取到群成员信息，放弃写入execl')


if __name__ == '__main__':
    options = Options()
    options.binary_location = r'C:/Program Files/Google/Chrome/Application/chrome.exe'
    res = login('群号')
    list_a = dispose(res)
    write_execl(list_a)