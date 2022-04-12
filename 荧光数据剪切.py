import os
import shutil
import time
time_y=time.localtime()[0]
print(time_y)
time_m=time.localtime()[1]

path1=r'G:\experiments'

path2=r'G:\荧光数据'
fil1=str(time_y)+'-'+str(time_m)
path2=os.path.join(path2,fil1)

if os.path.exists(path2)==False:
    os.mkdir(path2)

print(path2)
fileall=os.listdir(path1)
allfile=[]
for i in fileall:
    a=os.path.join(path1,i)
    shutil.move(a,path2)
