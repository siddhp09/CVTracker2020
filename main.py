
from bs4 import BeautifulSoup
import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import pandas as pd
import xlrd
import xlwt
import xlsxwriter
# Link: https://www.worldometers.info/coronavirus/#countries

service = Service('C:/Users/psidd/Downloads/chromedriver_win32 (1)/chromedriver.exe')
service.start()
driver = webdriver.Remote(service.service_url)
driver.get("https://www.worldometers.info/coronavirus/#countries")
time.sleep(20) # Let the user actually see something!

#driver = webdriver.Chrome("C:/Users/psidd/Downloads/chromedriver_win32 (1)/chromedriver.exe")
content = driver.page_source
soup = BeautifulSoup(content, features="html.parser")
names = []
total_Cases = []
table = soup.find('table', attrs={'class': 'table table-bordered table-hover main_table_countries dataTable no-footer'})
body = table.find('tbody')
row = body.findAll('tr')
index = 10

for x in range(index):
    row_single = row[x]
    tDiv = row_single.find('td')
    p = tDiv.find('a')
    name = str(row_single.find('td').find('a'))
    index = name.find('>')
    end = name.find('<', index, len(name))
    if name[index+1:end] != 'Non':
        names.append(name[index+1:end])

#index = 10

for x in row:
    tDiv = x.find('td')
    name = str(tDiv.string)
    if name != 'None' and name != 'China' and name != 'Italy' and name != 'Iran' and name!='Spain' and name!='Germany' and name!='USA' and name != 'France' and name != 'S. Korea' and name != 'UK':
        names.append(name.strip())

for y in row:
    cases = y.find('td', attrs={'class': 'sorting_1'})
    number = str(cases.string)
    comma = number.find(',')
    total_Cases.append(int(str(number[0:comma]) + str(number[comma+1:len(number)])))



# for a in soup.findAll('a', attrs={'class': 'mt_a'}):
#    names.append(a.text)
print(names)
print(total_Cases)
excel_file = 'CV Latest.xlsx'
df = pd.DataFrame({'Country':names,'Total Cases':total_Cases})
#df.to_excel(writer, sheet_name=sheet_name)
#workbook = writer.book
#worksheet = writer.sheets[sheet_name]

df.to_excel('CV Latest.xlsx', index=False, encoding='utf-8')
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
#df.to_csv('CV Update.csv', index=False, encoding='utf-8')

workbook = writer.book

chart = workbook.add_chart({'type': 'column'})


# Configure the series of the chart from the dataframe data.
chart.add_series({
   'values':     '=Sheet1!$B$2:$B$175',
   'gap':        2,
})

# Configure the chart axes.
chart.set_y_axis({'major_gridlines': {'visible': False}})
