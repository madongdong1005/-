import os

from selenium import webdriver
import time
from lxml import etree
from selenium.webdriver import ActionChains
import xlwt


class BaiDuMapSpider():

    def __init__(self, city_name, address_name, path="./tools/chromedriver.exe"):
        self.bro = webdriver.Chrome(executable_path=path)
        self.city_name = city_name
        self.address_name = address_name

    def run_action_chain(self):
        self.bro.get("https://map.baidu.com/@14103155,5724155,13z")
        but1 = self.bro.find_element_by_xpath("//div[@id='tool-container']/div[1]/div[1]/a[1]")
        but1.click()
        time.sleep(2)

        but2 = self.bro.find_element_by_id("selCityCityWd")
        but2.clear()
        but2.send_keys(self.city_name)

        but3 = self.bro.find_element_by_id("selCitySubmit")
        but3.click()

        time.sleep(2)
        but4 = self.bro.find_element_by_id("sole-input")
        but4.send_keys(self.address_name)

        but5 = self.bro.find_element_by_id("search-button")
        but5.click()
        time.sleep(2)

    def get_info(self):
        self.run_action_chain()
        while True:
            html = etree.HTML(self.bro.page_source)
            lis = html.xpath("//ul[@id='cards-level1']//ul[@class='poilist']/li")
            for li in lis:
                result_name = ''
                result_addr = ''
                result_tel = ''
                if li.xpath("./div[@class='cf']/div[3]/div[1]/span[1]/a/text()"):
                    result_name = li.xpath("./div[@class='cf']/div[3]/div[1]/span[1]/a/text()")[0].strip()
                if li.xpath("./div[@class='cf']/div[3]/div[2]/span[1]/@title"):
                    result_addr = li.xpath("./div[@class='cf']/div[3]/div[2]/span[1]/@title")[0].strip()
                if li.xpath("./div[@class='cf']/div[3]/div[3]/text()"):
                    result_tel = li.xpath("./div[@class='cf']/div[3]/div[3]/text()")[0].strip().split(":")[-1]
                yield {"城市": self.city_name,
                       "地点": self.address_name,
                       "结果名称": result_name,
                       "结果地点": result_addr,
                       "结果电话": result_tel}

            target = self.bro.find_element_by_id("poi_page")
            ActionChains(self.bro).move_to_element(target).perform()
            self.bro.execute_script("arguments[0].scrollIntoView();", target)
            time.sleep(2)
            if etree.HTML(self.bro.page_source).xpath("//div[@id='poi_page']/p/span[last()]/a/@class"):
                break
            self.bro.find_element_by_xpath("//div[@id='poi_page']/p/span[last()]/a").click()
            time.sleep(2)

    def write_into_excel(self):
        if not os.path.exists("./data"):
            os.mkdir("./data")
        wb = xlwt.Workbook()
        sh = wb.add_sheet("sheet1")
        for ind1, info1 in enumerate(self.get_info()):
            print(info1)
            row = sh.row(ind1)
            for ind2, info2 in enumerate(info1.items()):
                row.write(ind2, info2[1])

        wb.save(f"./data/{self.city_name}_{self.address_name}.xls")
        self.bro.close()
        self.bro.quit()


if __name__ == '__main__':
    baidu_map = BaiDuMapSpider("成都", "郫县团结修车")
    baidu_map.write_into_excel()