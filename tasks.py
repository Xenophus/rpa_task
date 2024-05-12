# coding = utf-8
# python3.10

import os
import datetime

import requests
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.by import By

# task 1
# 1.获取汇率中间价json数据
rate_url = 'https://www.amcm.gov.mo/api/v1.0/cms/financial_info?QueryType=ExchangeRate&Begin=20240510&End=20240510'
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome'
                         '/124.0.0.0 Safari/537.36'}
rate_json = requests.get(url=rate_url, headers=headers).json()
print(rate_json)

# 2.生成excel文件
file_name = '澳门银行同业汇率中间价_20240510.xlsx'
if os.path.exists(file_name):
    os.remove(file_name)

app = xw.App(visible=False, add_book=False)
wb_add = app.books.add()
wb_add.save(file_name)
wb_add.close()

# 3.向excel文件录入数据
rate_header = [['货币', '单位', '汇率中间价']]
rate_list = [[data['currency'], data['unit'], data['usdMeanValue']] for data in rate_json['data']]
rate_list = rate_header + rate_list

wb = app.books.open(file_name)
sht = wb.sheets["sheet1"]
sht.range('A1').value = rate_list
wb.save()

# task 2
# 读取excel表数据，删除第二列为100的行数据
shape = sht.used_range.shape
all_data = sht.range('A2:C%d' % shape[0]).value
real_data = [d for d in all_data if d[1] != 100]

# 添加备注
new_header = [['货币', '单位', '汇率中间价', '备注']]
new_data = []
for d in real_data:
    if d[2] > 1:
        d.append('True')
    elif d[2] < 1:
        d.append('False')
    else:
        d.append('')
    new_data.append(d)
sht.range('A1:C%d' % shape[0]).clear_contents()     # 清除原数据
sht.range('A1').value = new_header + new_data       # 重新写入添加备注后的数据
wb.save()

wb.close()
app.quit()

# task 3
'''
    selenium == 4.20.0
    chrome == 124
'''
# 1.搜索关键词 并获取标题和链接
driver = webdriver.Chrome()
driver.implicitly_wait(3)
driver.get("https://www.baidu.com")

kw = driver.find_element(By.ID, 'kw')
kw.send_keys('巴黎银行 洗钱')
su = driver.find_element(By.ID, 'su')
su.click()

ex = 0
cnt = 1
kw_list = []

while len(kw_list) < 10:
    a_xpath = '//div[@id="{}"]/div/div[1]/h3/a'.format(cnt)
    try:
        a = driver.find_element(By.XPATH, a_xpath)
        href = a.get_attribute('href')
        title = a.get_attribute('text')
        kw_list.append([title, href])
    except Exception:
        ex += 1
        if ex >= 3:
            next_page = driver.find_element(By.XPATH, '//div[@id="page"]/div/a[10]')    # 点击下一页
            next_page.click()
    finally:
        cnt += 1

# 2.生成 巴黎银行 洗钱 excel文件，将步骤1结果写入excel
date = datetime.datetime.now().strftime('%Y%m%d')
search_f = '巴黎银行 洗钱_%s.xlsx' % date
if os.path.exists(search_f):
    os.remove(search_f)

app_s = xw.App(visible=False, add_book=False)
wb_add_s = app_s.books.add()
wb_add_s.save(search_f)
wb_add_s.close()

search_header = [['标题', 'link']]

wb_s = app_s.books.open(search_f)
sht = wb_s.sheets["sheet1"]
sht.range('A1').value = search_header + kw_list
wb_s.save()

wb_s.close()
app_s.quit()
