#-----import modules--------------------------------------
import pandas as pd
import requests
import urllib.request
import urllib
import os
import time
from bs4 import BeautifulSoup
import csv
import re
import sys
import json
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Border

#-----functions--------------------------------------------
def lev(a, b):
    if a == "":
        return len(b)
    if b == "":
        return len(a)
    if a[-1] == b[-1]:
        cost = 0
    else:
        cost =1
    other = min([lev(a[:-1], b) + 1,
                 lev(a, b[:-1]) + 1,
                 lev(a[:-1], b[:-1]) + cost])
    return other

def length(a,b):
    length = len(a)+len(b)
    return length

def ratio(a,b):
    ratio = (1-round(lev(a,b)/length(a,b),3))*100
    return ratio

def sort(a):
    a = "".join(sorted(a.split()))
    return a
def trim(a):
    a = list(filter(lambda l: l!= '"',[l.strip() for l in a]))
    return a

def join(a):
    a = "".join(a.split())
    return a

def lower(a):
    a = "".join(trim(a)).lower()
    return a
#------get FTSE100 and Top 200 Law lists------------------

# for FTSE 100 index
pages = [1, 2, 3, 4, 5, 6]
output_file = r'\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\ftse_100_list.xlsx'
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
        row = list(filter(lambda r: r != '""', [tr.text.strip() for tr in td]))

        url_name = tr.find('a')
        link = r'http://www.londonstockexchange.com' + url_name.attrs['href']
        response = requests.get(link)
        soup = BeautifulSoup(response.text,"html.parser")
        t = soup.find('h1',{'class':'tesummary'})
        name = t.text.strip()

        if name != 'FTSE 100':
            ticker = re.search(r'.*\s', name)
            name = name.replace(ticker.group(),"").strip()
            row.append(name)

        while("" in row):
            row.remove("")

        l.append(row)
    lst.extend(l)
FTSE = [e for e in lst if e!=[]]
df=pd.DataFrame(FTSE,columns=['Code','Name','Cur','Price','+/-','%+/-','Full Name'])
df.to_csv(output_file, index=False, encoding = 'utf-8-sig')
df.to_excel(output_file, index=False)


# for top 200 law firms
output_file_2 = r'\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\law_firm_list.xlsx'
lst_2 = []
url = r'https://www.thelawyer.com/top-200-uk-law-firms/'
response = requests.get(url)
soup2 = BeautifulSoup(response.text, "html.parser")
table2 = soup2.find('tbody')
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
df2.to_csv(output_file_2, index=False, encoding='utf-8-sig') #avoid ()shown as funny characters
df2.to_excel(output_file_2, index=False)

#---for the report itself-----------------------------------------
#--------FTSE100--------------------------------------------------
scriptpath1 = r".\NameLibrary.py"
sys.path.append(os.path.abspath(scriptpath1))
scriptpath2 = r".\bin.py"
sys.path.append(os.path.abspath(scriptpath2))
with open(r".\NameLibrary.py", "r+") as f:
    NameDict = json.load(f)

with open(r".\bin.py","r+") as f2:
    Bin = json.load(f2)
# read excel file and put column to python list
file = r"\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\ftse_100_list.xlsx"
df = pd.read_excel(file, sheet_name=0)
mylist = df['Full Name'].tolist()
namelist = []
list2 = []
list3 = []

for n in mylist:
    match = re.search(r'PLC\s.*', n)
    if match:
        n = n.replace(match.group(),"")
        namelist.append(n)
    else:
        match = re.search(r'ORD\s.*',n)
        if match:
            n = n.replace(match.group(),"")
            namelist.append(n)
namelist = trim(namelist)

for nn in namelist:
    nn = nn.split()
    if nn[-1] in ['CO','AG','LD','GROUP','LTD','INTERNATIONAL','HOLDINGS','HLDGS']:
        nn = " ".join(nn[:-1])
        list2.append(nn)
    elif "".join(nn[-2:]) == "INVTST":
        nn = " ".join(nn[:-2])
        list2.append(nn)
    else:
        nn = " ".join(nn)
        list2.append(nn)

for i in list2:
    match = re.search(r'\([^()]*\)',i)
    if match:
        i = i.replace(match.group(),"")
        list3.append(i)
    else:
        match = re.search(r'GROUP\s.*',i)
        if match:
            i = i.replace(match.group(),"")
            list3.append(i)
        else:
            match = re.search(r'GROUP',i)
            if match:
                i= i.replace(match.group(),"")
                list3.append(i)
            else:
                list3.append(i)
list3 = [l.replace("&", "")for l in list3]
list3 = trim(list3)

wb = openpyxl.load_workbook(file)
s = wb.get_sheet_by_name('Sheet1')
row = 2
for nl in list3:
    s.cell(row, 8).value = nl
    row += 1
head = s.cell(1, 8)
head.value = "Clean Name"
head.font = Font(bold=True)
wb.save(file)

#compare with report file
report = r"\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\8. WORKINGS_Sep - Copy.xlsx"
sheetname = ['Library', 'PSL', 'Draft']
for sh in sheetname:
    df = pd.read_excel(report, sheet_name=sh)
    acct = df['accname'].tolist()
    #print(acct)
    matching = list(NameDict.keys())
    ftse = list(NameDict.values())
    wrongname = list(Bin.keys())
    ticker = list(Bin.values())

    wb2 = openpyxl.load_workbook(
        r'\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\8. WORKINGS_Sep - Copy.xlsx')
    Tab = wb2.get_sheet_by_name(sh)
    symbol = []

    for n in list2:
        nj = "".join(n.split())

        if n == nj:
            n = r'\b' + n + r'\s'
            for m in acct:
                match = re.search(n,m,flags=re.I)
                if match:

                    if m in matching:
                        mindex = acct.index(m)+1
                        Tab.cell(mindex+1, 11).value = nj
                        wb2.save(r'\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\8. WORKINGS_Sep - Copy.xlsx')

                    else:
                        if m in wrongname:
                            pass
                        else:
                            print(acct.index(m), match.group(), m)
                            user = input("do you want to add it into the dictionary? y/n: ")
                            if user != "n":
                                ftse.append(nj)
                                matching.append(m)
                                Tab.cell(mindex+1, 11).value = nj
                                wb2.save(r'\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\8. WORKINGS_Sep - Copy.xlsx')
                            else:
                                wrongname.append(m)
                                ticker.append(nj)

        else:
            symbol.append(n)
            for s in symbol:
                for m in acct:
                    match = re.search(n,m, flags = re.I)
                    if match:
                        mindex = acct.index(m)+1
                        Tab.cell(mindex+1, 11).value = s
                        wb2.save(r'\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\8. WORKINGS_Sep - Copy.xlsx')


 # above part misses royal shell and lloyds

    for acctname in acct:
        for a, b in NameDict.items():
            if acctname == a:
                tickersymbol = b
                acctindex = acct.index(acctname)+1
                if Tab.cell(acctindex+1, 11).value:
                    pass
                else:
                    Tab.cell(acctindex+1, 11).value = b
                    wb2.save(r'\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\8. WORKINGS_Sep - Copy.xlsx')


my_dict = dict()
for f, m in zip(ftse, matching):
    my_dict[m] = f
    dict.update(my_dict)

NameDict.update(my_dict)

with open(r'.\NameLibrary.py', 'w') as outfile:
    json.dump(NameDict, outfile)

bin_dict = dict()
for t, w in zip(ticker, wrongname):
    bin_dict[w] = t
    dict.update(bin_dict)

Bin.update(bin_dict)

with open(r".\bin.py", "w") as outfile2:
    json.dump(Bin, outfile2)

