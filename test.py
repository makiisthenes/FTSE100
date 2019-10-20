from functools import reduce
import sys
import os
import json
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Border
import re
import pandas as pd
scriptpath1 = r"./NameLibrary.py"
sys.path.append(os.path.abspath(scriptpath1))
scriptpath2 = r"./bin.py"
sys.path.append(os.path.abspath(scriptpath2))
with open(r"./NameLibrary.py", "r+") as f:
    NameDict = json.load(f)

with open(r"./bin.py","r+") as f2:
    Bin = json.load(f2)

def lev(a, b):
    if a == "":
        return len(b) # if a == "", then len(a) -> i = 0, while len(b) -> j; min(0,j) = 0, therefore lev(a,b) = max (0,j) = j

    if b == "":
        return len(a) # if b == "", then len(b) -> j = 0, while len(a) -> i; min(i,0) = 0, therefore lev(a,b) = max(i,0) = i

    if a[-1] == b[-1]:
        cost = 0  # a[-1] = ai, b[-1]=bj, if ai = bj, then deleting both final strings would not result in potential edit

    else:
        cost =1  # a[-1] = ai, b[-1]=bj, if ai <> bj, then deleting both final strings would result in potential one edit
                 # can assign any number as weight -> substitution can be more costy than deletion/insertation

    other = min([lev(a[:-1], b) + 1,  # A: a[:-1] -> string a with characters up till an-1; deleting a character itself has one edit

                 lev(a, b[:-1]) + 1, # B: b[:-1] -> string b with characters up till an-1; deleting a character itself has one edit

                 lev(a[:-1], b[:-1]) + cost])  # C  # if min(i,j) = 0, then lev(a,b) = max (a,b); otherwise lev(a,b)=min(A, B, C)


    #ratio = other/length

    return other
    #return ratio

def length(a,b):
    length = len(a)+len(b)
    return length

def ratio(a,b):
    ratio = (1-round(lev(a,b)/length(a,b),3))*100
    return ratio

def join(a):
    a = "".join(a.split())
    return a

def trim(a):
    a = list(filter(lambda l: l!= '"',[l.strip() for l in a]))
    return a
def lower(a):
    a = "".join(trim(a)).lower()
    return a
#----------------------------------------------------------------------------------------------
file = r"./ftse100_list.xlsx"
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
#wb.save(file)

#compare the files-------------------------------------------
report = r"./report.xlsx"
sheetname = ['Library', 'PSL', 'Draft']
for sh in sheetname:
    df = pd.read_excel(report, sheet_name=sh)
    acct = df['accname'].tolist()
#    matching = list(NameDict.keys())
 #   ftse = list(NameDict.values())
    wrongname = list(Bin.keys())
   # ticker = list(Bin.values())
    wb2 = openpyxl.load_workbook(report)
    Tab = wb2.get_sheet_by_name(sh)

    for n in list3:
        if n == join(n):
            nn = r'\b'+n+r'\s'
            for m in acct:
                match = re.search(nn,m,flags=re.I)
                if match and m not in wrongname:
                    mindex = acct.index(m)+2
                    Tab.cell(mindex,3).value = n
                    wb2.save(report)
        else:
            for m in acct:
                match = re.search(n,m,flags=re.I)
                if match and m not in wrongname:
                    mindex = acct.index(m)+2
                    Tab.cell(mindex,3).value = n
                    wb2.save(report)
                else:
                    pass

