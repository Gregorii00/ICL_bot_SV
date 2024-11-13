import datetime
from calendar import monthrange
import math

import pandas as pd
from openpyxl import load_workbook


spreadsheet_id = '1MQ0hNCotH-ou1jp3ODHeBhdww1w6UhG8X588-n-db6c'
url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet_id + '/export?format=xlsx&rtpof=true&sd=true'
sheet_name = 'График работы(Ноябрь1)'

df = pd.read_excel(url, sheet_name=sheet_name, engine='openpyxl', comment='#', header=None)
dw = df.columns.ravel()
df2= df
df2

now_year = datetime.datetime.now().year
now_month = datetime.datetime.now().month
days = monthrange(now_year, now_month)[1]
date_data = {}
for i in range(1, days + 1):
    date_data[i] = []
name = []
id_name = []
jj = 0
kk = 0
for i in df[dw[1]]:
    n = []
    if len(str(i).split(' ')) > 1 and jj >29:
        n.append(i)
        n.append(kk)
        n.append(kk+1)
        id_name.append(jj)
        id_name.append(jj+1)
        name.append(n)
        kk+=2
    jj+=1
data_col = {}
for j in range(3, len(dw)):
    k = 0
    n = []
    for i in df[j]:
        if k in id_name:
            n.append(i)
        k+=1
    data_col[j-2] = n
print(name)
print(data_col)
for i in range(len(name)):
    opers = {}
    k=0
    print(name[i])
    for j in data_col:
        print(j)
        opers_time = []
        if k == name[i][1] or k == name[i][2]:
            op = opers.get(name[i][0], [])
            print(data_col[j][k])
            op.append(data_col[j][k])
            opers[name[i][0]] = op
        k+=1
    print(opers)

