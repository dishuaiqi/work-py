import time
import pandas as pd
import os
from string import digits


files=[]
path1=r'D:\Users\Administrator\Desktop\送样单'
# file_path=r'D:\Users\Administrator\Desktop\送样单'

adress=os.listdir(path1)
for file in adress:
    files.append(os.path.join(path1,file))

# files = sorted(files, key=lambda x: os.path.getmtime(os.path.join(path1, x)))#格式解释:对files进行排序.x是files的元素,:后面的是排序的依据.   x只是文件名,所以要带上join.
files = sorted(files, key=lambda x: time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(os.path.getctime(x))), reverse=True)#格式解释:对files进行排序.x是files的元素,:后面的是排序的依据.   x只是文件名,所以要带上join.
files.reverse()
print(files)
allf=[]

for file1 in files:
    # print(file1)
    a=pd.read_excel(file1)
    if '阜阳禾丰' in file1:
        场区 = pd.read_excel(file1)['阜阳禾丰检测送样单']
        # print('123')
        场区=str(场区.tolist()[0]).replace(' ', '')[5:].split('送')[0]
        heade = a[(a['阜阳禾丰检测送样单'] == '样品编号')].index.tolist()
        样本信息 = pd.read_excel(file1, header=heade[0] + 1, dtype=object)

    else:
        安徽禾丰检测送样单=str(pd.read_excel(file1).columns.values[0])
        # print(安徽禾丰检测送样单)
        场区 = pd.read_excel(file1)[安徽禾丰检测送样单]

        if str(场区[0]).replace(' ', '')=='nan':
            场区 = str(场区[1]).replace(' ', '')[5:].split('送')[0]
        else:
            场区=str(场区[0]).replace(' ', '')[10:].split('送')[0]



        # 场区 = pd.read_excel(file1)['安徽禾丰检测送样单']
        heade = a[(a[安徽禾丰检测送样单] == '样品编号')].index.tolist()
        if len(heade) == 0:
            heade.append(7)
        样本信息 = pd.read_excel(file1, header=heade[0] + 1, dtype=object)
        样本信息 = 样本信息[~(样本信息['样品编号'].isnull())] #去空值
        样本信息.index = range(len(样本信息['样品编号']))
    if '部'  in str(场区):
        pass
    elif '场'  in str(场区):
        pass
    elif '站'  in str(场区):
        pass
    else:
        场区=str(场区)+'场'

    #     场区 = '工程部'
    if '生物安全' in str(场区) and "阜阳" not in str(场区) :
        场区='工程部'
    if "李相君"in file1:
        场区 = str(file1).split('(')[1][:-7]
        print(场区)
    if '马店区' in str(场区) :
        场区='马店育肥场'

    if str(场区)=='利辛服务部':
        if '利辛服务部' in file1:
            场区='利辛服务部'
        else:
            场区=str(file1).split('_')[1]
            场区=str(场区).split('送')[0]
    列总= 样本信息.columns.values
    列总2=''.join(列总)

    if '备注' not in 列总2:

        备注 = 样本信息[列总[2]]
    #     备注=样本信息['备注']
    # elif str(场区) == '工程部':
    #     备注=样本信息['备注']
    # elif '马店母猪场' == str(场区):
    #     备注 = 样本信息['备注']
    # elif '巩店公猪站' == str(场区):
    #     备注 = 样本信息['备注']
    else:

        备注 = 样本信息['备注']


    for i in 列总:
        if "类" in i:
            样品类型=样本信息[i]
    去除数字样本类型=[]
    for s in 样品类型:
        s=str(s)
        remove_digits = str.maketrans('', '', digits)
        res = s.translate(remove_digits)

        去除数字样本类型.append(res)
    样品类型=去除数字样本类型

    print(场区)

    a = ','.join('%s' % i for i in 样品类型)
    # print(a)
    a = a.replace('nan', '')
    样品类型 = a.split(',')
    混检索引2 = list(样本信息.loc[备注.str.contains('混', na=False)].index)
    混检索引1 = list(样本信息.loc[备注.str.contains('单', na=False)].index)
    # 混检索引1=list(样本信息.loc[备注.str.contains('非瘟',na=False)].index)

    混检索引 = 混检索引2 + 混检索引1


    混检索引.sort()

    if '送样时一定要低温运输' in str(样本信息['样品编号'].tolist()[-1]):  # 判断最后一行是否有备注
        样品编号 = 样本信息['样品编号'][:-1].tolist()
    elif '测非瘟病原' in  str(样本信息['样品编号'].tolist()[-1]):
        print(file1)
        样品编号 = 样本信息['样品编号'][:-2].tolist()
    else:
        样品编号 = 样本信息['样品编号'].tolist()
    if 样品编号[-1]=='测非瘟病原、非瘟抗体':
        样品编号=样品编号[:-1]

    样品 = []
    if str(场区) == '工程部' :  # 生物安全工程部
        样本信息 = 样本信息[~(样本信息['名字'].isnull())]
        样本信息.index = range(len(样本信息['名字']))
        # print(样本信息)
        # print(混检索引)
        名字=样本信息['名字'].tolist()
        场区1=样本信息['场区'].tolist()
        if  str(样本信息['名字'].tolist()[-1]):  # 判断最后一行是否有备注
            样品编号 = 样本信息['样品编号'][:-1].tolist()
        else:
            样品编号 = 样本信息['样品编号'].tolist()
        混检索引2 = list(样本信息.loc[备注.str.contains('混', na=False)].index)
        混检索引1 = list(样本信息.loc[备注.str.contains('单', na=False)].index)
        # 混检索引1=list(样本信息.loc[备注.str.contains('非瘟',na=False)].index)

        混检索引 = 混检索引2 + 混检索引1

        # print(混检索引)
        if len(混检索引) > 1:
            for i in range(len(混检索引) - 1):
                if str(名字[混检索引[i]]) == str(名字[混检索引[i+1] - 1]):  # 判定最后一个是否一样
                    样品.append(str('工程部') + str(场区1[混检索引[i]]) + str(样品类型[混检索引[i]]) + str(名字[混检索引[i]]))
                else:
                    样品.append(str('工程部') + str(场区1[混检索引[i]]) + str(样品类型[混检索引[i]]) + str(名字[混检索引[i]]) + '、' + str(名字[混检索引[i+1] - 1]))



            if str(名字[混检索引[-1]])==str(名字[-1]):
                样品.append(str('工程部') + str(场区1[混检索引[-1]]) + str(样品类型[混检索引[-1]]) + str(名字[-1]))
            else:
                样品.append(str('工程部') + str(场区1[混检索引[-1]]) + str(样品类型[混检索引[-1]]) + str(名字[混检索引[-1]]) + '、' + str(名字[-1]))
            # 样品.append(str('工程部'))
        else:
            if len(名字)>1:
                if str(名字[混检索引[-1]]) == str(名字[-1]):
                    样品.append(str('工程部') + str(场区1[混检索引[-1]]) + str(名字[混检索引[-1]]))
                else:
                    样品.append(str('工程部') + str(场区1[混检索引[-1]]) + str(名字[混检索引[-1]]) + str(名字[-1]))
            else:
                样品.append(str('工程部') + str(场区1[0]) + str(名字[混检索引[0]]))
        allf.extend(样品)

    elif '督查部' in file1 and '徐恩培' not in pd.read_excel(file1)['安徽禾丰检测送样单'][3] :


        列 = 样本信息.columns.values
        for i in 列:
            if "采样" in i :
                # print(i)
                采样环境 = 样本信息[i].tolist()
        # 采样环境 = 样本信息[列[3]].tolist()


        if len(混检索引) > 1:
            for i in range(len(混检索引) - 1):
                if str(样品编号[混检索引[i]]) == str(样品编号[混检索引[i + 1] - 1]):
                    样品.append(str(场区) + str(采样环境[混检索引[i]])+str(样品类型[混检索引[i]]) + str(样品编号[混检索引[i]]))
                else:
                    样品.append(str(场区) + str(采样环境[混检索引[i]])+str(样品类型[混检索引[i]]) + str(样品编号[混检索引[i]]) + '-' + str(样品编号[混检索引[i + 1] - 1]))

            if str(样品编号[-1]) == 'nan':
                样品.append(str(场区) +str(采样环境[混检索引[-1]])+ str(样品编号[混检索引[-1]]))
            # print(样本信息)
            # print(混检索引)
            # print(样品类型[混检索引[-1]])

            样品.append(str(场区) + str(采样环境[混检索引[-1]])+str(样品类型[混检索引[-1]]) + str(样品编号[混检索引[-1]]) + '-' + str(样品编号[-1]))

        else:
            if str(样品编号[0]) == str(样品编号[-1]):
                样品.append(str(场区) + str(采样环境[混检索引[0]])+str(样品类型[0]) + str(样品编号[0]))
            else:
                样品.append(str(场区) + str(采样环境[混检索引[0]])+str(样品类型[0]) + str(样品编号[0]) + '-' + str(样品编号[-1]))
            # 样品.append(str(场区) + str(样品类型[0]+str(样品编号[0])))
        allf.extend(样品)


    else:

        if len(混检索引) > 1:
            for i in range(len(混检索引) - 1):
                if str(样品编号[混检索引[i]]) == str(样品编号[混检索引[i + 1] - 1]):
                    样品.append(str(场区) + str(样品类型[混检索引[i]]) + str(样品编号[混检索引[i]]))
                else:
                    样品.append(str(场区) + str(样品类型[混检索引[i]]) + str(样品编号[混检索引[i]]) + '-' + str(样品编号[混检索引[i + 1] - 1]))

            if str(样品编号[-1]) == 'nan':
                样品.append(str(场区) + str(样品编号[混检索引[-1]]))


            样品.append(str(场区) + str(样品类型[混检索引[-1]]) + str(样品编号[混检索引[-1]]) + '-' + str(样品编号[-1]))

        else:
            if str(样品编号[0]) == str(样品编号[-1]):
                样品.append(str(场区) + str(样品类型[0]) + str(样品编号[0]))
            else:
                样品.append(str(场区) + str(样品类型[0]) + str(样品编号[0]) + '-' + str(样品编号[-1]))

        allf.extend(样品)

print(allf)
print(len(allf))
# print(样品编号)
样品信息表 = pd.DataFrame(allf)

aa = time.strftime("%Y-%m-%d", time.localtime())
path2=r'D:\Users\Administrator\Desktop'
path3=os.path.join(path2,aa + '样品信息.xlsx')
样品信息表.to_excel(path3)
print('over!')






