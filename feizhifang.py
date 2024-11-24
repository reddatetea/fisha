'''
根据产成品出入库流水账统计废纸房产品入库数量
'''
import os
import datetime
import easygui
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Border, Side, Fill, Alignment
import formatPainter
import re

dic = {'010201': '软抄',
 '010213': '定制产品',
 '010211': '定制产品',
 '010212': '莱特牛皮缝线',
 '010301': '锐意缝线本',
 '010311': '文稿纸',
 '010303': '锐意软抄',
 '010308': '锐意专利作业本',
 '010304': '锐意专利牛皮本',
 '010309': '锐意处理产品',
 '010310': '锐意无线胶装',
 '010501': '新锐缝线',
 '010503': '新锐软抄',
 '010504': '新锐防近视',
 '010599': '新锐定制（原木）',
 '010598': '新锐处理产品',
 '010606': '电商缝线',
 '010607': '电商抄本',
 '010608': '定制台板缝线',
 '0107': '外贸',
 '010602': '备课本',
 '010305': '锐意美术本',
 '010601': '电商美术本'}

col_lst = ['单据日期','业务类型','单据编号','单据类型','仓库','存货分类编码','存货分类','存货编码','存货','规格型号','入库数量',
 '入库数量2',
 '入库计量单位组合',
 '入库金额',
 '出库数量',
 '出库数量2']

# fname = r"F:\a00nutstore\008\zw08\废纸房\出入库流水9.26-10.25导出数据.xlsx"
msg = '请点选出入库流水'
fname = easygui.fileopenbox(msg)
path,_ = os.path.split(fname)
os.chdir(path)
df = pd.read_excel(fname, skiprows = 7,dtype= {'Unnamed: 6':str})
df = df.iloc[:, 1:]
df.columns = col_lst
df = df[~df['存货编码'].isnull()]
df1 = df[['存货分类编码','存货编码','存货','规格型号','入库数量2']]
df2 = df1.assign(feipin = (
 df1['存货分类编码']
    .str[:6]
    # .map(lambda x:x.zfill(6))
    .map(dic)
))

df3 = df2[~df2.feipin.isnull()]
result = pd.pivot_table(df3,index = 'feipin',values = '入库数量2',aggfunc = 'sum')
result.loc['合计','入库数量2'] = result['入库数量2'].sum()
dic1 = {v:k for k,v in dic.items()}
result = result.assign(biama = (
    result.index
    .map(dic1)
))
result = result.sort_values('biama')
result = result[['入库数量2']]
result.index.name = '产品大类'
result.columns =  ['单位（件）']
result.to_excel('废品入库数量.xlsx',index = True)
os.startfile('废品入库数量.xlsx')





