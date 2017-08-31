#!/usr/bin/env python
# -*- coding:utf-8 -*-

from bs4 import BeautifulSoup
from urllib import parse
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from multiprocessing import pool
from xlutils.copy import copy
import requests
import time
import re
import xlwt
import xlrd
import random
import sys
import socket

urls = []
pages = []
company = []
comurl = []
no = 0
headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.101 Safari/537.36',
    'connection': 'keep-alive',
    'Cookie': 'ssuid=204669008; TYCID=8bb8c0f08b1311e7a4bdb71d2798129c; uccid=34dcec360fe442c9db6bede3697e50f1; aliyungf_tc=AQAAAFU7p2SChQYAW9TtdOEluIWT1b3P; csrfToken=P4rxJqOyDtl4jhBqP8PWxOp5; _csrf=EB31vxDX+iQ7v1MTnmShvQ==; OA=k2LyAejDnkKESbi7TtIgXiIE+08ev66/UjrTSqq0G65KANMdEN7XkkcwsF/4lGZu; _csrf_bk=20d0d096d1baca1a4856a0909004168e; Hm_lvt_e92c8d65d92d534b0fc290df538b4758=1503830173; Hm_lpvt_e92c8d65d92d534b0fc290df538b4758=1503830424'
}


proxy_list = {
    '222.125.34.102:8080',
    '60.178.131.181:8081'
}


#proxy_ip = random.choice(proxy_list)
#proxies = {'http': proxy_ip}

f = xlwt.Workbook()
sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)


def CompanyName(companyurl):
    dcap = dict(DesiredCapabilities.PHANTOMJS)
    dcap["phantomjs.page.settings.userAgent"] = headers
    dcap["phantomjs.page.settings.loadImages"] = False
    driver = webdriver.PhantomJS(executable_path='/Users/MAC/phantomjs/phantomjs-2.1.1-macosx/bin/phantomjs', desired_capabilities=dcap)
    driver.set_window_size(1366, 768)

    # 设置用户名密码
    # driver.get("http://www.tianyancha.com/login")
    # driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[2]/input").send_keys("13818264698")
    # driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[3]/input").send_keys("qianjianan817")
    # driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[5]").click()
    # time.sleep((random.randint(2, 5)))

    #companyurl = 'https://sh.tianyancha.com/search/p2?key=%E8%BF%9B%E5%87%BA%E5%8F%A3%E8%B4%B8%E6%98%93'
    print(companyurl)
    try:
        driver.get(companyurl)
        time.sleep((random.randint(2, 5)))
        content = driver.page_source.encode('utf-8')
        driver.set_page_load_timeout(10)
        driver.set_script_timeout(10)
        driver.quit()
        soup = BeautifulSoup(content, 'lxml')
        print(soup)
    except Exception as e:
        print(e)
        pass

    try:
        name = soup.select('div > div > div > div.col-9.search-2017-2.pr10.pl0 > div.b-c-white.search_result_container > div > div.search_right_item > div.row.pb5 > div.col-xs-10.search_repadding2.f18 > a > span')
        comurl = soup.select('div > div > div > div.col-9.search-2017-2.pr10.pl0 > div.b-c-white.search_result_container > div > div.search_right_item > div.row.pb5 > div.col-xs-10.search_repadding2.f18 > a')
    except Exception as e:
        print(e)

    global no
    for url, companyname in zip(comurl, name):
        sheet1.write(no, 0, companyname.get_text())
        sheet1.write(no, 1, url.get('href'))
        no += 1
        f.save('公司2.xls')

        print(url.get('href'))
        print(companyname.get_text())


def get_information():
    address = []
    #url = 'https://www.tianyancha.com/company/420738689'
    url = []


    data = xlrd.open_workbook('公司2.xls')
    table = data.sheets()[0]
    nrows = table.nrows
    wb = copy(data)
    ws = wb.get_sheet(0)
    for i in range(0, 60):
        url = table.row_values(i)[1]
        print(url)
        time.sleep((random.randint(5, 10)))

        try:
            # 设置用户名密码
            # driver.get("http://www.tianyancha.com/login")
            # driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[2]/input").send_keys("13818264698")
            # driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[3]/input").send_keys("qianjianan817")
            # driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[5]").click()
            # time.sleep((random.randint(2, 5)))

            dcap = dict(DesiredCapabilities.PHANTOMJS)
            dcap["phantomjs.page.settings.userAgent"] = headers
            dcap["phantomjs.page.settings.loadImages"] = False
            driver = webdriver.PhantomJS(executable_path='/Users/MAC/phantomjs/phantomjs-2.1.1-macosx/bin/phantomjs', desired_capabilities=dcap)
            driver.set_window_size(1366, 768)
            driver.get(url)
            time.sleep((random.randint(2, 5)))
            content = driver.page_source.encode('utf-8')
            driver.set_page_load_timeout(10)
            driver.set_script_timeout(10)
            driver.close()
            driver.quit()
            soup = BeautifulSoup(content, 'lxml')

        except Exception as e:
            print(e)
            pass

        try:
            address = soup.select('#company_web_top > div.companyTitleBox55.pt20.pl30.pr30 > div.company_header_width.ie9Style > div > div > div > span.in-block.overflow-width.vertical-top.emailWidth')[1].get_text()
            telephone = soup.select('div.companyTitleBox55.pt20.pl30.pr30 > div.company_header_width.ie9Style > div > div.f14.new-c3.mt10 > div.in-block.vertical-top.overflow-width.mr20 > span:nth-of-type(2)')[0].get_text()
        except Exception as e:
            print(e)
            pass

        print(address)
        print(telephone)

        ws.write(i, 2, address)
        ws.write(i, 3, telephone)
        wb.save('公司2.xls')


if __name__ == "__main__":
    for i in range(6, 11):
        companyurl = 'https://sh.tianyancha.com/search/p{}?key=%E8%BF%9B%E5%87%BA%E5%8F%A3%E8%B4%B8%E6%98%93'.format(i)
        CompanyName(companyurl)


    #get_information()
