from bs4 import BeautifulSoup
import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import pandas as pd

service = Service('C:/Users/psidd/Downloads/chromedriver_win32 (1)/chromedriver.exe')
service.start()
driver = webdriver.Remote(service.service_url)
driver.get("https://www.worldometers.info/coronavirus/#countries")
time.sleep(20)

content = driver.page_source
soup = BeautifulSoup(content, features="html.parser")
names = []
total_Cases = []
table = soup.find('table', attrs={'class': 'table table-bordered table-hover main_table_countries dataTable no-footer'})
body = table.find('tbody')
row = body.findAll('tr')
index = 25

for x in range(index):
    row_single = row[x]
    tDiv = row_single.find('td')
    p = tDiv.find('a')
    name = str(row_single.find('td').find('a'))
    indexF = name.find('>')
    end = name.find('<', indexF, len(name))
    if x == 0:
        names.append("World")
    elif x == 23:
        names.append("Czechia")
    elif x == 24:
        names.append("Russia")
    elif name[indexF + 1:end] != 'Non':
        names.append(name[indexF + 1:end])

for x in range(len(row) - 25):
    row_single = row[x + 25]
    tDiv = row_single.find('td')
    name = str(tDiv.string)
    if name != 'None' and name != 'China' and name != 'Italy' and name != 'Iran' and name != 'Spain' and name != 'Germany' and name != 'USA' and name != 'France' and name != 'S. Korea' and name != 'UK':
        names.append(name.strip())

var = 0
for y in range(len(row)):
    row_single = row[y]
    cases = row_single.find('td', attrs={'class': 'sorting_1'})
    number = str(cases.string)
    comma = number.find(',')
    if var == 0:
        var = var + 1
        string1 = str(number[0:comma]) + str(number[comma + 1:len(number)])
        comma2 = string1.find(',')
        string2 = str(string1[0:comma2]) + str(string1[comma2 + 1:len(string1)])
        total_Cases.append(int(str(string2)))
    else:
        if comma != -1 and (y != 1 and y != 2 and y != 4 and y != 13 and y != 21 and y != 26 and y != 81) :
            total_Cases.append(int(str(number[0:comma]) + str(number[comma + 1:len(number)])))
        elif (y != 1 and y != 2 and y != 4 and y != 13 and y != 21 and y != 26 and y != 81) :
            total_Cases.append(int(str(number)))

excel_file = 'CV Latest.xlsx'
df = pd.DataFrame({'Country': names, 'Total Cases': total_Cases})
df.to_excel('CV Latest.xlsx', index=False, encoding='utf-8')
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')

