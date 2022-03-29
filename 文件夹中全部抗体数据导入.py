import openpyxl
import pandas as pd
import numpy as np
import os
import time
wb=openpyxl.load_workbook(r'D:\Users\Administrator\Desktop\抗体数据.xlsx')
ws=wb.active
path1=r'D:\Users\Administrator\Desktop\抗体数据'
file1=os.listdir(path1)


fileall=[]
for i in file1:
    s=os.path.join(path1,i)

    fileall.append(s)


count蓝耳=[] #统计多少个蓝耳
countgb=[]
countge=[]
count猪瘟=[]
count口蹄疫=[]
count非瘟=[]
count圆环=[]
count细小=[]
for i in fileall:
    if '蓝耳' in i:
        count蓝耳.append(i)
        # x=23
        列=2
    elif 'gb' in i:
        countgb.append(i)
        # x=23
        列=4
    elif 'ge' in i :
        countge.append(i)
        # x=22
        列=5
    elif '猪瘟' in i:
        count猪瘟.append(i)
        # x = 22
        列 = 3
    elif '口蹄疫'in i:
        count口蹄疫.append(i)
        # x = 22
        列 = 6
    elif '非'in i:
        count非瘟.append(i)
        # x = 22
        列 = 8
    elif '圆环' in i:
        count圆环.append(i)
        # x = 22
        列 = 7
    elif '细小' in i:
        count细小.append(i)
        列=9

# print(count蓝耳)
s=[]
countall=[count圆环,count非瘟,count口蹄疫,count猪瘟,countge,countgb,count蓝耳,count细小]
sa=[]
dir={'蓝耳':[],'猪瘟':[],'gb':[],'ge':[],'口蹄疫':[],'圆环':[],'非瘟':[],'细小':[]}
liename1=['蓝耳','猪瘟','gb','ge','口蹄疫','圆环','非瘟','细小']
for i in countall:
    print(i)
    for j in i:
        if '蓝耳' in j:
            x=23
        elif 'gb'in j:
            x=23
        elif '圆环' in j:
            x = 23
        elif '细小' in j:
            x=23
        else:
            x=22
    if len(i) ==0:
        pass
    if len(i)==1:
        数据1 = open(i[0])

        数据1 = np.array(数据1.read().split()[x:])  # 蓝耳的写入
        # print(len(数据1))
        数据1.shape = 8, 13
        数据 = 数据1.flatten('F')[8:]

        lieming=str(i[0]).split('\\')[-1]
        lieming=str(lieming).split('.')[0][8:]
        dir[lieming]=数据.tolist()
    if len(i)==2:
        数据1 = open(i[0])
        数据1 = np.array(数据1.read().split()[x:])  # 蓝耳的写入
        # print(数据1)
        数据1.shape = 8, 13
        数据1 = 数据1.flatten('F')[8:]

        数据2 = open(i[1])
        数据2 = np.array(数据2.read().split()[x:])  # 蓝耳的写入
        数据2.shape = 8, 13
        数据2 = 数据2.flatten('F')[8:]
        数据=数据1.tolist()+数据2.tolist()
        lieming = str(i[0]).split('\\')[-1]
        lieming = str(lieming).split('.')[0][8:][:-1]
        print(lieming)
        dir[lieming] = 数据
    if len(i)==3:
        数据1 = open(i[0])
        数据1 = np.array(数据1.read().split()[x:])  # 蓝耳的写入
        # print(数据1)
        数据1.shape = 8, 13
        数据1 = 数据1.flatten('F')[8:]

        数据2 = open(i[1])
        数据2 = np.array(数据2.read().split()[x:])  # 蓝耳的写入
        数据2.shape = 8, 13
        数据2 = 数据2.flatten('F')[8:]


        数据3 = open(i[2])
        数据3 = np.array(数据3.read().split()[x:])  # 蓝耳的写入
        数据3.shape = 8, 13
        数据3 = 数据3.flatten('F')[8:]
        数据=数据1.tolist()+数据2.tolist()+数据3.tolist()
        lieming = str(i[0]).split('\\')[-1]
        lieming = str(lieming).split('.')[0][8:][:-1]
        dir[lieming] = 数据


df = pd.DataFrame.from_dict(dir, orient='index')
df=df.transpose()#行列转换
df.loc['阴性对照1']=''
df.loc['阴性对照2']=''
df.loc['阳性对照1']=''
df.loc['阳性对照2']=''
print(df)
df.index=np.arange(1,len(df)+1)
aa = time.strftime("%Y-%m-%d", time.localtime())
path2=r'D:\Users\Administrator\Desktop'
path3=os.path.join(path2,aa + '抗体数据.xlsx')
df.to_excel(r'D:\Users\Administrator\Desktop\抗体数据.xlsx')
df.to_excel(path3)
