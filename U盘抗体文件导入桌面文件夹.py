import os
import shutil
import time
time_y=time.localtime()[0]
print(time_y)
time_m=time.localtime()[1]

path1=r'G:\FC'

path2=r'D:\Users\Administrator\Desktop\抗体数据'

path3=r'D:\Users\Administrator\Desktop'
fil1=str(time_m)+'月抗体数据'
path3=os.path.join(path3,fil1)

if os.path.exists(path3)==False:
    os.mkdir(path3)

print(path3)
fileall=os.listdir(path2)
# 先把path2转入path3
allfile=[]
for i in fileall:
    a=os.path.join(path2,i)
    shutil.move(a,path3)

#path1转入path2
fil2=os.listdir(path1)
allpath1=[]
for i in fil2:
    a=os.path.join(path1,i)
    shutil.move(a,path2)