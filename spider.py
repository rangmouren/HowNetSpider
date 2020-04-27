import random
import time

import xlsxwriter
import xlwt
from lxml import etree
from selenium import webdriver
# from selenium.webdriver.opera.options import Options
from selenium.webdriver.chrome.options import Options
from config import *
from excel import generate_excel,read_xlrd


class Z_W_Spider(object):

    @staticmethod
    def get_html(url):
        driver = webdriver.Chrome(executable_path=Driver_Path)
        driver.maximize_window()
        driver.get(url)
        time.sleep(4)
        return driver

    @staticmethod
    def operating_html(driver, name, company):
        driver.find_element_by_id('au_1_sel').click()
        time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id="au_1_sel"]/option[2]').click()
        time.sleep(1)
        driver.find_element_by_id('au_1_value1').send_keys(name)
        time.sleep(1)
        if company:
            driver.find_element_by_id('au_1_value2').send_keys(company)
        time.sleep(1)
        try:
            driver.execute_script("SingleSearchOnclick('&ua=1.21')")
        except Exception as  e:
            print(e)
        time.sleep(5)
        iframe = driver.find_element_by_id('iframeResult')
        driver.switch_to_frame(iframe)
        page = driver.find_element_by_class_name('pagerTitleCell').text
        page = int(page[4:-5]) // 20
        html = etree.HTML(driver.page_source)
        datas = Z_W_Spider.get_data(html)
        time.sleep(random.uniform(2, 4))
        for i in range(page):
            page_next = driver.find_element_by_id('Page_next')
            page_next.click()
            time.sleep(random.uniform(2, 4))
            html = etree.HTML(driver.page_source)
            datas.extend(Z_W_Spider.get_data(html))
        driver.quit()
        return datas

    @staticmethod
    def get_data(html):
        datas = []
        trs = html.xpath('//tr[@bgcolor]')
        for tr in trs:
            title = tr.xpath('./td//a[@class="fz14"]/text()')[0]
            url = tr.xpath('./td[2]/a/@href')[0][55:]
            authors = tr.xpath('./td[@class="author_flag"]/a[@class="KnowledgeNetLink"]//text()')
            authors = "|".join(authors)
            try:
                source = tr.xpath('./td/a[@target="cdmdNavi"]/font/text()')[0]
            except:
                source = tr.xpath('./td/a[@target="cdmdNavi"]/text()')[0]
            degree = tr.xpath('./td[@align="center"]/text()')[0].strip()
            time = tr.xpath('./td[@align="center"]/text()')[1].strip()
            data = {
                "title": title,
                'url': 'https://kns.cnki.net/KCMS/detail/detail.aspx?dbcode=CMFD&' + url,
                "authors": authors,
                "source": source,
                "degree": degree,
                "time": time,
            }

            datas.append(data)
        return datas

    @staticmethod
    def run(url, name, company):
        dirver = Z_W_Spider.get_html(url)
        time.sleep(2)
        return Z_W_Spider.operating_html(dirver, name, company)


if __name__ == '__main__':
    # for i in read_xlrd(excelFile='./机械运载学部2019版.xlsx'):
        name = '田红旗'
        company = '中南大学'
        z_w_url = 'https://kns.cnki.net/kns/brief/result.aspx?dbprefix=CDMD'
        try:
            result=Z_W_Spider.run(z_w_url, name,company)
            if result:
                generate_excel(result, name)
        except Exception as e:
            print('写入失败，失败原因：{}'.format(e))

