import pandas as pd
import requests
import urllib.request
import urllib
import os
import time
from bs4 import BeautifulSoup 
import pprint
# for FTSE 100 index
pages = [1, 2, 3, 4, 5, 6]
output_file = r'C:\Users\chenyx\Documents\Evelyn\Practise\Python learning\FTSE100\ftse_list.csv'
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


# for top 200 law firms
output_file_2 = r'C:\Users\chenyx\Documents\Evelyn\Practise\Python learning\FTSE100\law_firm_list.csv'
lst_2 = []
url = r'https://www.thelawyer.com/top-200-uk-law-firms/'
response = requests.get(url)
soup2 = BeautifulSoup(response.text, "html.parser")
table2 = soup2.find('table', {'class': 't1'})
table_rows2 = table2.findAll('tr')
a = []
for tr2 in table_rows2:
    td2 = tr2.findAll('td')
    row2 = [tr2.text.strip() for tr2 in td2]
    while("" in row2):
        row2.remove("")
    a.append(row2)
law = [x for x in a if x != []]
head = law[0]
del law[0]
df2 = pd.DataFrame(law,columns=head)
df2.to_csv(output_file_2, index=False, encoding='utf-8-sig') #avoid ()showing funny characters






