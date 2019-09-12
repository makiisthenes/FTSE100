import pandas as pd
import requests
import urllib.request
import urllib
import os
import time
from bs4 import BeautifulSoup 
import pprint
pages = [1, 2, 3, 4, 5, 6]
output_file = r'C:\Users\chenyx\Documents\Evelyn\Practise\Python learning\FTSE100 Report\list.csv'
lst=[]
for page in pages:
    url = r'https://www.londonstockexchange.com/exchange/prices-and-markets/stocks/indices/summary/summary-indices-constituents.html?index=UKX&page={}'.format(
        page)
    response = requests.get(url)
    soup = BeautifulSoup(response.text,"html.parser")
    table = soup.find('table',{'class':'table_dati'})
    table_rows = table.findAll('tr')
    l=[]
    for tr in table_rows:
        td=tr.findAll('td')
        row = [tr.text.strip() for tr in td]
        while("" in row):
            row.remove("")
        
        l.append(row) 
    lst.extend(l)
FTSE = [e for e in lst if e!=[]]
df=pd.DataFrame(FTSE,columns=['Code','Name','Cur','Price','+/-','%+/-'])
df.to_csv(output_file, index=False)





