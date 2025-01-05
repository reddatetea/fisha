'''
1. 将拼多多收入与订单合并
2. 对收入分别按订单汇总、按存货编码汇总
202405-202409
保留中间过程
'''

import easygui
import os
import pandas as pd
import re
import numpy as np
import easygui

regax = r'（订单(?P<dingdanhao>\d{6}-\d{15})）'
pattern = re.compile(regax)
# mat = pattern.search(string)
def getDingdanhao(s):
    mat = pattern.search(s)
    if mat :
        s = mat.group('dingdanhao')
    else :
        s = ''
    return s

#所有收入
def chuliFnameShouru(fname):
    # fname_shouru = r"F:\a00nutstore\008\zw08\电商\拼多多\pdd-mall-bill-detail(100306881)_2024-09-30-07-09-49_3.csv"
    df_shouru0 = pd.read_csv(fname,header = 4,encoding = 'gb18030')
    df_shouru1 = df_shouru0[['商户订单号','发生时间','收入金额（+元）','支出金额（-元）','账务类型','备注','业务描述']]
    df_shouru1 = df_shouru1.dropna(thresh = 2)   #至少有三个缺失值 
    df_shouru1['商户订单号']= df_shouru1['商户订单号'].mask(df_shouru1['商户订单号'].isnull(),df_shouru1['备注'].map(getDingdanhao))
    df_shouru1['商户订单号'] = df_shouru1['商户订单号'].replace('','888888-888888888888888')
    df_shouru1.columns = ['dingdanhao', 'riqi', 'ru', 'chu', 'leixing', 'beizhu', 'miaoshu']
    df_shouru2 = df_shouru1.loc[df_shouru1.leixing != '提现']
    df_shouru2 = df_shouru2.assign(shouru = df_shouru2['ru'] + df_shouru2['chu'])
    return df_shouru2

dic = {'五': 5,
       '六': 6,
       '七': 7,
       '八': 8,
       '九': 9,
       '十': 10,
       '四': 4,
       '三': 3,
       '二': 2,
       '一': 1}


def addBen(string):
    # string = r'(81本)'
    pattern1 = r'(?P<ben>\d+)本'
    pattern2 = r'(?P<ben>[十九八七六五四三二一])本'
    regexp1 = re.compile(pattern1)
    mat1 = regexp1.search(string)
    regexp2 = re.compile(pattern2)
    mat2 = regexp2.search(string)
    if mat1:
        ben = int(mat1.group('ben'))
    else:
        if mat2:
            ben = dic.get(mat2.group('ben'))
        else:
            ben = 1
    return ben


#所有订单
path_dingdan = r"F:\a00nutstore\008\zw08\电商\拼多多\拼多多订单创业至当下"
data_dingdan = []
for i in os.listdir(path_dingdan):
    j = os.path.join(path_dingdan,i)
    if  os.path.isfile(j) :
        df = pd.read_csv(j)
        data_dingdan.append(df)
    else :
        continue


def readDingdan(fname) :
    df_dingdan = pd.read_csv(fname)
    return df_dingdan

def chuliDingdan(df_dingdan0):
    df_dingdan1 = df_dingdan0[['订单号','发货时间','商家编码-规格维度','商家实收金额(元)','商品数量(件)','商品规格']]
    df_dingdan1.columns = ['dingdanhao','riqi','code','jiner','bao','guige']
    df_dingdan1[df_dingdan1['dingdanhao'].isnull()]
    df_dingdan1['hanliang'] = df_dingdan1['guige'].map(addBen)
    df_dingdan1 = df_dingdan1.assign(ben=df_dingdan1['bao'] * df_dingdan1['hanliang'])
    print(df_dingdan1.columns.to_list())
    df_dingdan2 = df_dingdan1[['dingdanhao','riqi','code','jiner','bao','guige','ben']]
    df_dingdan2 = df_dingdan2.set_index('dingdanhao')
    return df_dingdan2


df_dingdan0 = pd.concat(data_dingdan)
df_dingdan2 = chuliDingdan(df_dingdan0)
        

#拼多多收入
choice = easygui.choicebox('是计算当月收入还是累计收入',choices =['当月','累计'])
if  choice == '当月':
    qijian = easygui.enterbox('请输入期间，"2024-12"')
    fname_shouru = easygui.fileopenbox('请点选当月收入文件')
    df_xiaoshou0 = chuliFnameShouru(fname_shouru)
else :
    path = easygui.fileopenbox('请点选各月收入文件夹')
    data = []
    for i in os.listdir(path):
        j = os.path.join(path,i)
        if  os.path.isfile(j):
            df = chuliFnameShouru(j)
            data.append(df)
        else :
            continue
    df_xiaoshou0 = pd.concat(data)
    
shouru_pivot =df_xiaoshou0.pivot_table(index = 'dingdanhao',values =['ru','chu','shouru'],aggfunc = np.sum)
result0 =pd.merge(shouru_pivot,df_dingdan2,left_index = True,right_index = True,how = 'left')
#空值填充为0
for i in ['chu','ru','shouru','jiner','bao','ben']:
    result0[i] = result0[i].fillna(0)
#code填充
result0['code'] = result0['code'].replace('\t',None)
result0['guige'] = result0['guige'].replace('\t',None)
result0['code'] = result0['code'].fillna('A8888')
result0['guige'] = result0['guige'].fillna('A8888')

#日期填充
result0['riqi'] = result0['riqi'].replace('\t',None)
result0['riqi'] = result0['riqi'].ffill()
result0['riqi'] = result0['riqi'].bfill()

gp = result0.groupby('dingdanhao')
#将订单总金额按各产品的明细销售额进行分配
data = []
for k,v in gp:
    # print(k,v)
    mean = v.shouru.mean()
    total = v['jiner'].sum()
    # print('mean',mean)
    # print('total',total)
    v['result_jiner'] = v.shouru*(v.jiner)/total
    data.append(v)
result1 = pd.concat(data)
result1 = result1.reset_index()
result1 = result1.rename({'riqi': '日期',
 'dingdanhao': '订单号',
 'ru': '收入',
 'chu': '支出',
 'shouru': '净收入',
 'code': '存货编码',
 'jiner': '订单金额',
 'bao': '包',
   'ben': '本',
   'guige': '商品规格',
                          
 'result_jiner': '最终金额'},axis = 1)
result1  = result1[['日期', '订单号', '收入', '支出', '净收入', '存货编码', '订单金额', '包','本', '最终金额','商品规格']]
result1 ['最终金额']= np.where(result1 ['最终金额'].isnull(),result1 ['净收入'],result1['最终金额'])
total = result1.sum().T.to_frame().T
total.日期  = '合计'
total.订单号  = ''
total.存货编码  = ''
total.商品规格  = ''
result2 = pd.concat([result1,total])
result2.to_excel(f'拼多多{qijian}净收入明细.xlsx',index = False)


# df_xiaoshou0.to_excel('拼多多收入未关联订单.xlsx',index = False)




