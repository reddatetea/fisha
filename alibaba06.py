import numpy as np
import pandas as pd
import os
import easygui
import openpyxl
import re



def readexcel(fname):
    df = pd.read_excel(fname,dtype = {'订单号':'str','Offer_ID':'str', 'SKU_ID':'str'})
    return df

def alibabaXiaoshouMingxi(fname):
    df = pd.read_excel(fname,dtype = {'订单号':'str','Offer_ID':'str', 'SKU_ID':'str'})
    col1s = ['订单号','卖家公司名称', '卖家会员名','订单状态', '订单创建时间', '订单付款时间', '收货人姓名', '收货地址', '邮编',
       '联系手机',  '货运公司', '运单号']
    col2s = ['货品总价', '运费（元）', '商家改价（元）',
           '实付款（元）','单价']
    for col in col1s:
        df[col] =  df[col].ffill()
    for col in col2s :
        df[col] = df[col].fillna(0)
    return df

def gethuohao(df0,df1):
    df1 = df1[~df1['订单号'].isin(df0['订单号'].to_list())]
    df = pd.concat([df0,df1])
    return df
    
def getZhuangtai(df):
    # df = df[(df['订单状态'] == '已收货') | (df['订单状态'] == '交易成功')]
    df = df[df['订单状态'] == '交易成功']
    return df

fname0 = r"F:\a00nutstore\008\zw08\电商\阿里巴巴1688\阿里巴巴订单中没有品名(1).xlsx"
# fname = r"F:\a00nutstore\008\zw08\电商\阿里巴巴1688\阿里巴巴202240726-20241031.xlsx"
qijian = easygui.enterbox('请输入期间，"2024-12"')
fname = easygui.fileopenbox('请点选阿里巴巴销售明细文件')
df0 = readexcel(fname0)
df1 = alibabaXiaoshouMingxi(fname)
df = gethuohao(df0,df1)
df = getZhuangtai(df)
df_dingdan = df[['订单号','单品货号','数量','货品总价','运费（元）','商家改价（元）','实付款（元）','单价']]
df_dingdan['jiner'] = df_dingdan['数量']*df_dingdan['单价']
pivot = df_dingdan.pivot_table(index = '单品货号',values = ['数量','货品总价','运费（元）','商家改价（元）','实付款（元）','jiner'] ,aggfunc = 'sum')

#支付宝收入
# fname_shouru = r"F:\a00nutstore\008\zw08\电商\阿里巴巴1688\2088340986968423-20241105-141006887-收入.csv"
fname_shouru = easygui.fileopenbox(f'请点选阿里巴巴资金交易查询文件,期间为{qijian}')
df_shouru = pd.read_csv(fname_shouru,header = 2,encoding = 'gb18030',dtype = {'商户订单号':'str','业务基础订单号':'str', '支付宝交易号':'str','支付宝流水号':'str',})
df_shouru = df_shouru[~df_shouru['支付宝交易号'].isnull()]
df_shouru = df_shouru[['入账时间', '支付宝交易号', '支付宝流水号', '商户订单号', '账务类型', '收入（+元）', '支出（-元）',
             '备注', '业务基础订单号', '业务订单号']]
df_shouru['收入（+元）'] = df_shouru['收入（+元）'].fillna(0)
df_shouru['收入（+元）'] = df_shouru['收入（+元）'].replace(' ',0)
df_shouru['支出（-元）'] = df_shouru['支出（-元）'].fillna(0)
df_shouru['支出（-元）'] = df_shouru['支出（-元）'].replace(' ',0)
df_shouru1 = df_shouru[df_shouru['账务类型'].isin(['其它', '转账', '在线支付', '分账', '退款', '退款（交易退款）']) ]
df_shouru1 = df_shouru1.astype({'入账时间':'datetime64[ns]'})
# df_shouru1 = df_shouru[~(df_shouru['账务类型'] == '充值（大额充值）')]
df_shouru1['收入（+元）'] = df_shouru1['收入（+元）'].astype('float64') 
df_shouru1['支出（-元）'] = df_shouru1['支出（-元）'].astype('float64') 
df_shouru1['jinshouru'] = df_shouru1['收入（+元）'] - df_shouru1['支出（-元）']
regex1 = re.compile(r'^1688.*_(\d+)')  
regex2 = re.compile(r'^1688.*-订单号(\d+)')
regex3 = re.compile(r'^1688.*[:\]](\d+)-.*')
regex4 = re.compile(r'^聚宝盆.*订单号:(\d+).*')
regex5 = re.compile(r'^代扣款.*订单号(\d+).*')
# df_shouru1.to_excel('shouru1-13.xlsx',index = False)
# string = r'T50060NP2268563451685578186'
# mat = re.search(regex1,string)
# mat
def getDingdanhao(str1,str2):
    if str1 ==  None:
        return str1
    else :
        try:
            mat = re.search(regex1,str2)
            return mat.group(1)
        except:
             try:
                mat = re.search(regex2,str2)
                return mat.group(1)
             except:
                 try:
                    mat = re.search(regex3,str2)
                    return mat.group(1)
                 except:
                     try:
                        mat = re.search(regex4,str2)
                        return mat.group(1)
                     except:
                         try:
                            mat = re.search(regex5,str2)
                            return mat.group(1)
                         except:
                             return str1
                         
               
  
        
df_shouru2 =  df_shouru1.assign(dingdanhao = df_shouru1.apply(lambda x:getDingdanhao(x['业务基础订单号'],x['备注']),axis = 1))
# df_shouru2
df_shouru2['dingdanhao'] = df_shouru2['dingdanhao'].fillna('8888888888888888888')
df_shouru2['dingdanhao'] = df_shouru2['dingdanhao'].astype('str')
df_shouru2['dingdanhao'] = df_shouru2['dingdanhao'].replace(' ','8888888888888888888')
# df_shouru2
pivot_shouru = pd.pivot_table(df_shouru2,index = 'dingdanhao',values = ['收入（+元）','支出（-元）','jinshouru'] ,aggfunc = 'sum')
pivot_shouru = pivot_shouru.reset_index()
result = pd.merge(pivot_shouru,df_dingdan,left_on = 'dingdanhao',right_on = '订单号',how = 'left')
result['订单号'] = result['订单号'].fillna('8888888888888888888')
result['单品货号'] = result['单品货号'].fillna('A8888')
result = result.assign(jiner = np.where(result['单品货号'] == 'A8888',result['jinshouru'],result.jiner))
# result.to_excel('shouru_dingdan分配前-1.xlsx',index = False)
gp = result.groupby('订单号')
#将订单总金额按各产品的明细销售额进行分配
data = []
for k,v in gp:
    # print(k,v)
    mean = v['jinshouru'].mean()
    total = v['jiner'].sum()
    # print('mean',mean)
    # print('total',total)
    v['result_jiner'] = v['jinshouru']*(v.jiner)/total
    data.append(v)
result1 = pd.concat(data)
# result1 = result1.reset_index()

result2 = result1[['dingdanhao', 'jinshouru',  '单品货号', '数量','单价', 'jiner', 'result_jiner']]
result2.columns = ['订单号', '净收入',  '单品货号', '数量','单价',  '金额', '最终金额']
result2 = result2.assign(最终金额 = np.where(result2['单品货号'] == 'A8888',result2['净收入'],result2['最终金额']))
# pivot = pivot.reset_index()
total =result2.sum().to_frame().T
result2 = pd.concat([result2,total])
result2.iloc[-1,0] = '小计'
result2.iloc[-1,1] = ''
result2.iloc[-1,2] = ''

result2.to_excel(f'阿里巴巴销售明细表按品名汇总{qijian}.xlsx',index = False)




