import pandas as pd
import numpy as np

df = pd.read_excel('PackingListFormat.xlsx', usecols=[0,1,2,3,4,5], header=10)
df.drop(df.tail(2).index,inplace=True)

forCsv = df.groupby(['Item Code'], as_index=False)['Item Code','Number of Pieces','Number of Cartons'].sum()
forCsv['Item Code'] = forCsv['Item Code'].astype(str)
forCsv['Number of Pieces'] = forCsv['Number of Pieces'].astype(int)

forCsv['Number of Cartons'] = forCsv['Number of Cartons'].astype(int)
forCsv.to_csv('packing.csv',index=False)
outDf = pd.read_csv('output.csv', names =['Item Code','Number of Pieces','Number of Cartons'], dtype=None, index_col =False)

outDf = outDf.groupby(['Item Code'], as_index=False)['Number of Pieces','Number of Cartons'].sum()

outDf['Item Code'] = outDf['Item Code'].astype(str)

outDf['Number of Pieces'] = outDf['Number of Pieces'].astype(int)

outDf['Number of Cartons'] = outDf['Number of Cartons'].astype(int)

common = forCsv.merge(outDf,on=['Item Code','Number of Pieces','Number of Cartons'])

print('Uncommon or wrong entries in Packing List are : ' )
print(forCsv[(~forCsv['Item Code'].isin(common['Item Code']))])

print('Uncommon or wrong entries in POITM screen are: ')
print(outDf[(~outDf['Item Code'].isin(common['Item Code']))])


