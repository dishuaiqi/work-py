import os
import shutil
import time
time_m=time.localtime()[1]

path1=r'D:\Users\Administrator\Desktop\送样单'

path2=r'D:\Users\Administrator\Desktop'
fil1=str(time_m)+'月送样单'
path2=os.path.join(path2,fil1)
if os.path.exists(path2)==False:
    os.mkdir(path2)
print(path2)
fileall=os.listdir(r'D:\Users\Administrator\Desktop\送样单')
allfile=[]
for i in fileall:
    a=os.path.join(path1,i)
    shutil.move(a,path2)
