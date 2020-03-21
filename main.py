from bs4 import BeautifulSoup
import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import pandas as pd

# Link: https://www.worldometers.info/coronavirus/#countries

service = Service('C:/Users/psidd/Downloads/chromedriver_win32 (1)/chromedriver.exe')
service.start()
driver = webdriver.Remote(service.service_url)
driver.get("https://www.worldometers.info/coronavirus/#countries")
time.sleep(20)  # Let the user actually see something!
content = driver.page_source
soup = BeautifulSoup(content, features="html.parser")
names = []
total_Cases = []
table = soup.find('table', attrs={'class': 'table table-bordered table-hover main_table_countries dataTable no-footer'})
body = table.find('tbody')
row = body.findAll('tr')
index = 20

for x in range(index):
    row_single = row[x]
    tDiv = row_single.find('td')
    p = tDiv.find('a')
    name = str(row_single.find('td').find('a'))
    indexF = name.find('>')
    end = name.find('<', indexF, len(name))
    if name[indexF + 1:end] != 'Non':
        names.append(name[indexF + 1:end])

for x in range(len(row) - 20):
    row_single = row[x + 20]
    tDiv = row_single.find('td')
    name = str(tDiv.string)
    if name != 'None' and name != 'China' and name != 'Italy' and name != 'Iran' and name != 'Spain' and name != 'Germany' and name != 'USA' and name != 'France' and name != 'S. Korea' and name != 'UK':
        names.append(name.strip())

for y in row:
    cases = y.find('td', attrs={'class': 'sorting_1'})
    number = str(cases.string)
    comma = number.find(',')
    if comma != -1:
        total_Cases.append(int(str(number[0:comma]) + str(number[comma + 1:len(number)])))
    else:
        total_Cases.append(int(str(number)))

excel_file = 'CV Latest.xlsx'
df = pd.DataFrame({'Country': names, 'Total Cases': total_Cases})
df.to_excel('CV Latest.xlsx', index=False, encoding='utf-8')


