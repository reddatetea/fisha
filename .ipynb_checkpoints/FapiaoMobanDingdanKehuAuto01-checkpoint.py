'''
tcloud中用销售订单生成税局发票模板
1. 分别按每个销售订单汇总，再按存货编码汇总，批量生成发票模板，存于同一文件夹下
2. 分别按每个客户汇总，再按存货编码汇总，批量生成发票模板，存于同一文件夹下
3. 选择税率
4. 折扣处理

全部自动生成
'''
import os
import re
import easygui
import openpyxl
import numpy as np
import pandas as pd
import shutil


def res_path(relative_path):
    """获取资源绝对路径"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

#读取发票模板
fname_fapiao = res_path('img\piao.xlsx')
df_fapiao = pd.read_excel(fname_fapiao,header = 2,dtype = {'商品和服务税收分类编码':'str'})
#wb0 = openpyxl.load_workbook(res_path('img/leiji.xlsx'))

def chuliZekou(string):
    if string == "运费":
        string = '折扣'
        return string
    else:
        for i in string:
            if i.isdigit():
                continue
            else:
                return string

        if len(set(list([*string]))) == 1:
            string = '折扣'
        else:
            string = string

    return string


def getPivot(df):
    bianba_fenlei = dict(zip(df['存货编码'],df['存货分类']))
    bianba_mingchen = dict(zip(df['存货编码'],df['存货名称']))
    bianba_daima = dict(zip(df['存货编码'],df['存货代码']))
    bianba_hanlian = dict(zip(df['存货编码'],df['件含量']))
    pivot = df.pivot_table(index = '存货编码',values = ['数量','数量（件）','含税金额'] ,aggfunc = 'sum')
    pivot = pivot.reset_index()
    pivot = pivot.assign(fenlei = pivot['存货编码'].map(bianba_fenlei))
    pivot = pivot.assign(mingchen = pivot['存货编码'].map(bianba_mingchen))
    pivot = pivot.assign(daima = pivot['存货编码'].map(bianba_daima))
    pivot = pivot.assign(hanliang = pivot['存货编码'].map(bianba_hanlian))
    dic = dict(zip(['fenlei','mingchen','daima','hanliang'],['存货分类','存货名称','存货代码','件含量']))
    pivot = pivot.rename(columns =  dic)
    guige_qian = []
    guige_hou = []
    for i in pivot['存货名称'].to_list():
        qian0,hou0 = guige(i)
        guige_qian.append(qian0)
        guige_hou.append(hou0)
    pivot = pivot.assign(qian = guige_qian)
    pivot = pivot.assign(hou = guige_hou)
    return pivot

def getFapiaoBen(d,shuilu,shuliang_fangsi):
    
    d['项目名称'] = d['存货编码'] +  d['hou']
    d['项目名称'] = d['项目名称'].str.replace('运费-','运费')
    # d['项目名称'] = d['项目名称'].str.split('-').str[0]
    d['商品和服务税收分类编码'] = '1060202010000000000'
    d['规格型号'] = d['qian']
    if shuliang_fangsi == '本数':
        d['单位'] = '本'
        d['商品数量'] = d['数量']
    else :
        d['单位'] = '件'
        d['商品数量'] = d['数量（件）']
        
    d['商品单价'] = ''
    d['金额'] = d['含税金额']
    d['税率'] = shuilu
    d['折扣金额'] = ''
    d['优惠政策类型'] = ''
    d = d[['项目名称',
     '商品和服务税收分类编码',
     '规格型号',
     '单位',
     '商品数量',
     '商品单价',
     '金额',
     '税率',
     '折扣金额',
     '优惠政策类型',]]
    return d



def getFapiaoMoban(gongsi,shuliang_fangsi):
    filename=''.join(['发票模板-',gongsi,'-',shuliang_fangsi,'.xlsx'])
    newname = os.path.join(path1,filename)
    shutil.copyfile(fname_fapiao, newname)
    return newname

def getFapiaoMobanDingdanhao(gongsi,dingdanhao,shuliang_fangsi):
    filename=''.join(['发票模板-',gongsi,'-',dingdanhao,'-',shuliang_fangsi,'.xlsx'])
    newname = os.path.join(path1,filename)
    shutil.copyfile(fname_fapiao, newname)
    return newname

def fengefu(string):
    num = len(string.split('-'))
    if num  == 3 :
        string = '-'.join([string.split('-')[0],string.split('-')[1]])
    elif num == 2 :
        if len(string.split('-')[0]) <= 4:
            string = string
        else :
            string = string.split('-')[0]
    else :
        string = string
    return string

def guige(string):
    if ('型' in string) and ('页' in string):
        qian0 = string.split('型')[0] + '型'
        hou0 =  string.split('型')[1] 
    elif ('型' in string) or ('页' in string):
        if '型' in string :
            qian0 = string.split('型')[0] + '型'
            hou0 =  string.split('型')[1] 
        else :
            qian0 = string.split('页')[0] + '页'
            hou0 =  string.split('页')[1] 
    else :
        qian0 = ''
        hou0 = ''
    return qian0,hou0    
            
# path = r"F:\repos\fisha\莱新销售订单0826-0925"
path = easygui.diropenbox('请点选销售订单所在文件夹')
# os.chdir(path)
#莱新销售订单超5000条，不能一次导出，分三次导出，并分别存于同一文件下，先将它们合并
data = []
for i in os.listdir(path):
    j = os.path.join(path,i)
    if  os.path.isfile(j):
        df = pd.read_excel(j)
        data.append(df)
    else :
        continue


df_xiaoshou0 = pd.concat(data)
df_xiaoshou1 = df_xiaoshou0.loc[df_xiaoshou0['单据执行状态'] != '合计']
lst1 = ['单据编号',
  '单据日期',
 '含税总金额',
 '存货名称',
 '存货分类',
 '存货编码',
  '存货代码',
  '数量',
  '件含量',
 '数量（件）',
  '含税单价',
 '含税金额',
        '客户']
df_xiaoshou2 = df_xiaoshou1[lst1]
df_xiaoshou2['存货编码'] = df_xiaoshou2['存货编码'].apply(lambda x:chuliZekou(x))
df_xiaoshou3 = df_xiaoshou2.loc[~df_xiaoshou2['单据编号'].isnull()]
df_xiaoshou3['存货代码'] = df_xiaoshou3['存货代码'].ffill()
s = []
for i in df_xiaoshou3['存货编码'].to_list():
    j = fengefu(i)
    s.append(j)
df_xiaoshou3['存货编码'] = s


#选择税率
shuilu = easygui.choicebox(msg = '请选择税率',choices = [0.13,0.01,0.02])
shuilu  = float(shuilu)
#选择数量的开具方式，件数or本数
shuliang_fangsi = easygui.choicebox(msg = '请选择数量开具方式',choices = ['件数','本数'])
# #选择发票开具方式
# fapiao_fangsi = easygui.choicebox(msg = '请选择发票开具方式',choices = ['按客户和存货编码','按销售订单和存货编码','按客户汇总'])





#按照发票开具方式，及数量选择方式，生成对应的文件夹
path1 = f'发票模板-按客户和存货编码-{shuliang_fangsi}'
path1 = os.path.join(path,path1)
if not os.path.exists(path1):
    try:
        os.mkdir(path1)
    except:
        pass

gp = df_xiaoshou3.groupby('客户')
for k,v  in gp:
    gongsi = k
    newname = getFapiaoMoban(gongsi,shuliang_fangsi)
    pivot = getPivot(v)
    # pivot1 = chuliMingchen(pivot)
    fapiao = getFapiaoBen(pivot,shuilu,shuliang_fangsi)
    with pd.ExcelWriter(newname, engine='openpyxl',mode='a', if_sheet_exists='overlay')  as writer:
        fapiao.to_excel(writer, sheet_name = '1-明细模板',startrow=3, header = False,index = False)

#按照发票开具方式，及数量选择方式，生成对应的文件夹
path1 = f'发票模板-按销售订单和存货编码-{shuliang_fangsi}'
path1 = os.path.join(path,path1)
if not os.path.exists(path1):
    try:
        os.mkdir(path1)
    except:
        pass

gp = df_xiaoshou3.groupby('单据编号')
for k,v  in gp:
    dingdanhao = k
    gongsi = v['客户'].to_list()[0]
    newname = getFapiaoMobanDingdanhao(gongsi,dingdanhao,shuliang_fangsi)
    # print(newname)
    pivot = getPivot(v)
    # pivot1 = chuliMingchen(pivot)
    fapiao = getFapiaoBen(pivot,shuilu,shuliang_fangsi)
    with pd.ExcelWriter(newname, engine='openpyxl',mode='a', if_sheet_exists='overlay')  as writer:
        fapiao.to_excel(writer, sheet_name = '1-明细模板',startrow=3, header = False,index = False)

#按照发票开具方式，及数量选择方式，生成对应的文件夹
path1 = f'发票模板-按客户汇总-{shuliang_fangsi}'
path1 = os.path.join(path,path1)
if not os.path.exists(path1):
    try:
        os.mkdir(path1)
    except:
        pass
     
gp = df_xiaoshou3.groupby('客户')
for k,v  in gp:
    gongsi = k
    newname = getFapiaoMoban(gongsi,shuliang_fangsi)
    v1 = v.sum().T.to_frame().T
    v1.loc[0,'存货名称'] = '本册'
    v1.loc[0,'存货分类'] = '本册'
    v1.loc[0,'存货编码'] = '本册'
    v1.loc[0,'存货代码'] = '本册'
    pivot = getPivot(v1)
    # pivot1 = chuliMingchen(pivot)
    fapiao = getFapiaoBen(pivot,shuilu,shuliang_fangsi)
    with pd.ExcelWriter(newname, engine='openpyxl',mode='a', if_sheet_exists='overlay')  as writer:
        fapiao.to_excel(writer, sheet_name = '1-明细模板',startrow=3, header = False,index = False)


    
    
    
    
    
    
    
    
    
    
    

