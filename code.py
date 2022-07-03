import os
import pandas as pd
inPath = os.path.join(".","input/")
outPath =os.path.join(".","output/")

get = input("檔案名稱:")  #寶齡富錦(1760)
dbPath = inPath + get +".xlsx"
data= pd.read_excel(dbPath)

KD = []
K = data['K(9,3)']
D = data['D(9,3)']

KD.append('NaN')
x=0

for i in range(1,len(data)):
    x = K[i-1] - D[i-1]
    y = K[i] - D[i]
    if x > 0 and y < 0:
        KD.append('死亡交叉')
    elif x < 0 and y > 0:
        KD.append('黃金交叉')
    else:
        KD.append('NaN')

data['KD'] = KD

data2 = data.loc[(data['KD']=='死亡交叉') | (data['KD']=='黃金交叉')]
data = data2['時間'].astype(str)
data2['時間'] = data
data2 = data2.reset_index()
data2 =data2.drop(columns=['index'])

SMA60 = data2['SMA60']
close = data2['收盤價']
KD = data2['KD']

BUY = []
for i in range(len(data2)):
    x = close[i] - SMA60[i]
    if x > 0 and KD[i]=='黃金交叉':
        BUY.append('買')
    elif x < 0 and KD[i]=='死亡交叉':
        BUY.append('賣')
    else:
        BUY.append('再考慮')

dict1 = {"日期": data2['時間'],  
        "K值": data2['K(9,3)'],
        "D值": data2['D(9,3)'],
        "黃金/死亡交叉" : data2['KD'],
        "RSI":data2['RSI 6'],
        "當天收盤股價": data2['收盤價'],
        "季線" : data2['SMA60'],
        "買或賣" : BUY
       }
df=pd.DataFrame(dict1)
print(df)
df.to_excel(outPath+get+".xlsx",encoding="utf8")