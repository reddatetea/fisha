'''
计算聚水潭中各平台销售
文件夹下各文件的命名应准确
拼多多、京东、京东荣佳、阿里巴巴、淘宝、天猫、分销商、京东商城
全平台（不包括分销部）

'''

import pandas as pd
import numpy as np
import easygui
import openpyxl
import os
import re

lst = ['图片',
 '商品编码',
 '款式编码',
 '商品标签',
 '供应商',
 '供应商款号',
 '供应商商品编码',
 '颜色规格',
 '商品简称',
 '商品名称',
 '商品品牌',
 '产品分类',
 '虚拟分类',
 '成本价',
 '销售数量',
 '价格为零的商品数量',
 '实发数量',
 '实发金额',
 '销售金额',
 '销售成本',
 '实发成本',
 '销售毛利',
 '销售毛利率',
 '销售均价',
 '退货数量',
 '实退数量',
 '退货金额',
 '退货成本',
 '实退成本',
 '实退金额',
 '退货毛利',
 '净销量',
 '净销售额',
 '净销售成本',
 '净销售毛利',
 '基本金额',
 '已付金额',
 '优惠金额',
 '运费收入',
 '运费支出',
 '净毛利率',
 '其它价格1',
 '其它价格2',
 '其它价格3',
 '其它价格4',
 '其它价格5',
 '其它属性1',
 '其它属性2',
 '其它属性3',
 '其它属性4',
 '其它属性5',
 '基本售价']
lst1 = [ '商品编码',
 '商品名称',
 '净销量',
 '净销售额',
 '净销售成本',
 '净销售毛利',
]
names = ['shangpinbianma','name','xiaoshouliang','xiaoshoue','xiaoshouchengben','maoli']
dic = dict(zip(lst1,names))

# fname = r"F:\a00nutstore\008\zw08\电商\聚水潭\各平台总销售.xlsx"
path = easygui.diropenbox('请点选聚水潭各平台销售所在文件夹')
# os.chdir(path)
# path = r'F:\a00nutstore\008\zw08\电商\聚水潭\202412'
lst_file = [i for i in os.listdir(path) if  (i.startswith('~') == False) and (i.endswith('.xlsx') == True)] # and i.startswith('~') == Fasle


def getXiaoshou(file):
    df = pd.read_excel(os.path.join(path,file),dtype =  {'商品编码':str,'其它属性4':str}) 
    df = df[~df['商品编码'].isnull()]
    return df 

dfs = []
pingtais = []
pingtai_df  = {}
pattern = re.compile(r'(?<=-)(\w+)(?=\.xlsx)')
for file in lst_file:

    mat = pattern.search(file)
    df = getXiaoshou(file)
    
    
    if mat :
        pingtai = mat.group(1)
        if pingtai == '全平台':
            df_quanpingtai = df
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)
        
           

        elif pingtai == '拼多多':
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)
  
        elif pingtai == '阿里巴巴':
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)
           
        elif pingtai == '淘宝':
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)
        elif pingtai == '天猫':
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)
        elif pingtai == '抖音':
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)
 
        elif pingtai == '京东':
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)
  
        elif pingtai == '京东商城':
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)
         
        elif pingtai == '京东荣佳':
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)
 
        elif pingtai == '分销商':
            pingtais.append(pingtai)
            pingtai_df[pingtai] = df
            dfs.append(df)

        else :
            continue
    else :
        continue
    

#双佳平台
pingtais_shuang = [i for i in pingtais if (i != '京东荣佳') and (i != '全平台')]
#双佳平台数据
dfs_shuangjia = [pingtai_df.get(i) for i in pingtais]
zhong = []
for pingtai in pingtais:
    
    df = pingtai_df.get(pingtai)
    # print(pingtai,df)
    df.insert(0,'pingtai',pingtai)
    # print(df)
    zhong.append(df)
    df = pd.DataFrame()
df_zhong = pd.concat(zhong)

#请输入期间
qijian = easygui.enterbox('请输入期间，"2024-12"')
df_zhong.insert(0,'期间',qijian)

#是否附加到原始数据中
isno = easygui.boolbox('是否将本期销售数据添加到原始数据中去')
if isno:
    fname = r"F:\a00nutstore\008\zw08\电商\聚水潭\聚水潭各平台总销售.xlsx"
    df = pd.read_excel(fname,sheet_name = '原始数据',dtype =  {'商品编码':str,'其它属性4':str})
    max_row = df.shape[0]
    with pd.ExcelWriter(fname, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df_zhong.to_excel(writer, sheet_name='原始数据', startrow=max_row + 1, header=False, index=False)

else :
    df_zhong.to_excel(f'各平台总销售{qijian}.xlsx',index = False)
os.startfile(fname)







