import sys
import os
import json
import openpyxl
from openpyxl import workbook, worksheet
import pandas as pd


filepath = r'\\Galileo\Public\Legal Intelligence\Customer Segmentation\BA\Ad Hoc Reports & Requests\2019\201909 - September\DAI-2093 - Kenneth Ume - Market Product Penetration Data Request - REPORT\checking.xlsx'
wb = openpyxl.load_workbook(filepath)
nomore = wb.get_sheet_by_name('nomore')
elist = []
for cell in nomore['A']:
    elist.append(cell.value)
    if cell.value is None:
        break
#print(elist)



scriptpath = r".\lawlibrary.py"
sys.path.append(os.path.abspath(scriptpath))
with open (r".\lawlibrary.py", "r+") as f:
    Dict = json.load(f)



kl = []
vl = []
for key, value in Dict.items():
    if value in elist: 
        pass
    else:

        if value in ['Cripps','Pemberton Greenish']:
            value = 'Cripps Pemberton Greenish'
            #print([key,value])
        if value in ['Ince & Co' ,'Gordon Dadds']:
            value = 'Ince'
            #print ([key,value])
        if value in ['Pitmans' ,'BDB']:
            value = 'BDB Pitmans'
            #print([key, value])
        if value in ['Bond Dickinson','Womble']:
            value = 'Womble Bond Dickinson'
        # print([key, value])
        if value in ['Berwin Leighton Paisner','Bryan Cave Leighton Paisner']:
            value == 'BCLP'
        
        else:
            pass

        kl.append(key)
        vl.append(value)

new_dic = dict()
for key, value in zip (kl, vl):
    new_dic[key] = value
    dict.update(new_dic)

Dict.update(new_dic) 
with open(r'.\lawlibrary.py', 'w') as outfile:
   json.dump(Dict, outfile)
#print(Dict)




    


