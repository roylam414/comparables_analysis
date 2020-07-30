from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import re
import os.path

pd.options.display.float_format = '{:,.2f}'.format

columns = ['stock_code','col_name', 'col_last', 'col_issued_shares','market_cap', 'PEratio','earnings', 'col_summary']
df = pd.DataFrame(index=None, columns = columns)

dirname = os.path.dirname(__file__)
fullpath = os.path.join(dirname, '.\stock_lists.xlsx')
df2 = pd.read_excel(fullpath)
stock_lists = df2['stock_lists']


chrome_options = Options()
chrome_options.add_argument('--headless')

for i in stock_lists:

    driver = webdriver.Chrome(options=chrome_options)
    r = driver.get("https://www.hkex.com.hk/Market-Data/Securities-Prices/Equities/Equities-Quote?sym={}&sc_lang=en".format(i))
    content = driver.page_source
    soup = BeautifulSoup(content,'html.parser')

    col_name = soup.find('p', class_='col_name').get_text()

    col_last = soup.find('span', class_='col_last').get_text()
    col_last = float(col_last)
    
    col_issued_shares = soup.find('span', class_='col_issued_shares').get_text()
    col_issued_shares = col_issued_shares.split()[0]
    col_issued_shares = float(re.sub(',','',col_issued_shares))
    
    market_cap = col_last * col_issued_shares
    
    PEratio = soup.find('dt', class_='ico_data col_pe').get_text()
    if PEratio == '-':
        PEratio = PEratio.replace('-','N/A')
        earnings = 'N/A'
        
       
    else:
        PEratio = float(PEratio.replace('x',''))
        earnings = (col_last/PEratio)*col_issued_shares
    
    

    col_summary = soup.find('div', class_='company_txt col_summary').get_text()
   
    lists = [[i,col_name, col_last, col_issued_shares,market_cap,PEratio, earnings, col_summary]]
    df_temp = pd.DataFrame(lists, index=None, columns = columns)
    df = df.append(df_temp, ignore_index=True)
    driver.close()

df = df.sort_values(by='market_cap', ascending=False)

with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    print(df[['col_name', 'col_last','market_cap', 'PEratio']])

#df.to_excel('.\comparables.xlsx',sheet_name='Sheet1',index=False)