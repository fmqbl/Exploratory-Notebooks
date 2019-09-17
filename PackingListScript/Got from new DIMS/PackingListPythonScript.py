import pandas as pd
import numpy as np

import math


def my_round(n, ndigits):
    part = n * 10 ** ndigits
    delta = part - int(part)
    # always round "away from 0"
    if delta >= 0.5 or -0.5 < delta <= 0:
        part = math.ceil(part)
    else:
        part = math.floor(part)
    return part / (10 ** ndigits)

def piecesDiff(row):
    if row['B'] == 'Both':
        p = row['Number of Pieces_x'] - row['Number of Pieces_y']
    else:
        p = 0 
    return abs(p)


def cartonsDiff(row):
    if row['B'] == 'Both':
        p = row['Number of Cartons_x'] - row['Number of Cartons_y']
    else:
        p = 0 
    return abs(p)

def grossDiff(row):
    if row['B'] == 'Both':
        p = row['Gross Weight_x'] - row['Gross Weight_y']
    else:
        p = 0 
    return abs(p)

def volumeDiff(row):
    if row['B'] == 'Both':
        p = row['Total Volumn CBM_x'] - row['Total Volumn CBM_y']
    else:
        p = 0 
    return abs(p)

df = pd.read_excel("PackingList903241.xlsx", usecols=[0,1,2,3,4,5], header=10)

df.drop(df.tail(2).index,inplace=True)

df.dropna(how='all', inplace=True)

forCsv = df.groupby(['Item Code'], as_index=False)['Item Code','Number of Pieces','Number of Cartons','Gross Weight','Total Volumn CBM'].sum()

forCsv['Item Code'] = forCsv['Item Code'].astype(str)

forCsv['Number of Pieces'] = forCsv['Number of Pieces'].astype('Int64')


forCsv['Number of Cartons'] = forCsv['Number of Cartons'].astype('Int64')

forCsv['Gross Weight'] = forCsv['Gross Weight'].astype(float)
forCsv['Total Volumn CBM'] = forCsv['Total Volumn CBM'].astype(float)

forCsv['Gross Weight'] = forCsv['Gross Weight'].apply(lambda x: '{:.2f}'.format(x))

forCsv['Total Volumn CBM'] = forCsv['Total Volumn CBM'].apply(lambda x: '{:.2f}'.format(x))

outDf = pd.read_csv("output.csv", names =['Item Code','Number of Pieces','Number of Cartons','Gross Weight','Total Volumn CBM'], dtype=object, index_col =False)

outDf = outDf.groupby(['Item Code'], as_index=False)['Number of Pieces','Number of Cartons', 'Gross Weight','Total Volumn CBM'].sum()

outDf['Item Code'] = outDf['Item Code'].astype(str)
outDf['Item Code'] = outDf['Item Code'].apply(lambda x : x.split('/')[0])
outDf['Number of Pieces'] = outDf['Number of Pieces'].astype('Int64')

outDf['Number of Cartons'] = outDf['Number of Cartons'].astype('Int64')
outDf['Gross Weight'] = outDf['Gross Weight'].astype(float)
outDf['Total Volumn CBM'] = outDf['Total Volumn CBM'].astype(float)

outDf['Gross Weight'] = outDf['Gross Weight'].apply(lambda x : '{:.2f}'.format(x))
outDf['Total Volumn CBM'] = outDf['Total Volumn CBM'].apply(lambda x : my_round(x,2))

outDf['Total Volumn CBM'] = outDf['Total Volumn CBM'].apply(lambda x : '{:.2f}'.format(x))

common = forCsv.merge(outDf,on=['Item Code','Number of Pieces','Number of Cartons','Gross Weight','Total Volumn CBM'])
packingList = forCsv[(~forCsv['Item Code'].isin(common['Item Code']))]#&(~forCsv['Number of Pieces'].isin(common['Number of Pieces'])) & (~forCsv['Number of Cartons'].isin(common['Number of Cartons']))]
poitm = outDf[(~outDf['Item Code'].isin(common['Item Code']))]#&(~outDf['Number of Pieces'].isin(common['Number of Pieces'])) & (~outDf['Number of Cartons'].isin(common['Number of Cartons']))]


m = {'left_only': 'PackingList', 'right_only': 'POITM', 'both': 'Both'}

result = packingList.merge(poitm, on=['Item Code'], how='outer', indicator='B')
#result = result.fillna(-1).astype(mainDict)
result['B'] = result['B'].map(m)

result.fillna(0, inplace=True)
result['Number of Pieces_y'] = result['Number of Pieces_y'].astype('Int64')
result['Number of Cartons_y'] = result['Number of Cartons_y'].astype('Int64')
result["Gross Weight_x"] = pd.to_numeric(result["Gross Weight_x"])
result["Gross Weight_y"] = pd.to_numeric(result["Gross Weight_y"])
result["Total Volumn CBM_x"] = pd.to_numeric(result["Total Volumn CBM_x"])

result["Total Volumn CBM_y"] = pd.to_numeric(result["Total Volumn CBM_y"])

result['Pieces Difference'] = result.apply(piecesDiff, axis=1)

result['Cartons Difference'] = result.apply(cartonsDiff, axis=1)

result['Gross Weight Difference'] = result.apply(grossDiff, axis=1)

result['Volume Difference'] = result.apply(volumeDiff, axis=1)

result.to_csv('ResultedFile.csv',index=False, header=['Item Code','Pieces in PL','Cartons in PL','GW in PL','Volumn in PL','Pieces in ETMS','Cartons in ETMS','GW in ETMS','Volumn in ETMS','Availability','Pieces Diff','Cartons Diff','Weight Diff','Volume Diff'])

























