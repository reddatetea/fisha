'''
计算电商绩效
'''
import re
import os
import openpyxl
import easygui
import numpy as np
import pandas as pd

mingxi_pinyin = {'收入': 'shouru',
 '成本': 'chengben',
 '打包': 'dabao',
 '耗材': 'haocai',
 '运费': 'yunfei',
 '软件': 'luanjian',
 '外包': 'waibao',
 '推广费': 'tuiguang',
 '工资': 'gongzi',
 '库管': 'kuguan',
 '办公': 'bangong',
 '其他': 'qita'}
column_pinyin = {'销售额': 'shouru',
 '调拨成本': 'chengben',
 '打包成本': 'dabao',
 '包装耗材': 'haocai',
 '运费': 'yunfei',
 '服务费': 'luanjian',
 '客服费': 'waibao',
 '平台推广': 'tuiguang',
 '工资': 'gongzi',
 '库管': 'kuguan',
 '办公': 'bangong',
 '其他': 'qita'
}
pinyin_column = {v:k for k,v in column_pinyin.items()}
xiangmu_pingtai = {'拼多多': '拼多多',
 '阿里巴巴': '阿里巴巴',
 '淘宝': '淘宝',
 '京东京喜京麦': '京东京喜',
 '京东商城': '京东商城',
 '抖音': '抖音',
 '聚水潭': '分销商'}
dingdan_jixiao = {'1688': '阿里巴巴',
 '{线下}': '线下',
 '京东': '京东商城',
 '分销商': '分销商',
 '惊喜': '京东京喜',
 '抖音': '抖音',
 '拼多多': '拼多多',
 '淘宝': '淘宝'}
pingtai_people = {'拼多多': '邵劭',
 '阿里巴巴': '刘宗威',
 '淘宝': '邵劭',
 '京东京喜': '刘宗威',
 '京东商城': '邵劭',
 '抖音': '杨杰',
 '分销商': '刘宗威'}
#先从用友T3畅易通系统导出本期间的收入、成本、销售费用明细（项目科目总账）
# fname = r"F:\a00nutstore\008\zww08\002电商\电商考核绩效电商销售分析\电商项目202501-03.xls"
fname = easygui.fileopenbox('请点选本期电商项目文件')
dfs = pd.read_excel(fname,sheet_name = None)

#电商项目各收入、费用明细
def GetMingxi(v):
    v.dropna(inplace = True)
    mingxi = dict(v.set_index('电商')['本期借方发生'])
    return mingxi

# for k,v in dfs.items():
#     if k == '收入':
#         shouru = GetMingxi(v)
#     elif k == '成本':
#         chengben = GetMingxi(v)
#     elif k == '打包':
#         dabao = GetMingxi(v)
#     elif k == '耗材':
#         haocai = GetMingxi(v)
#     elif k == '运费':
#         yunfei = GetMingxi(v)
#     elif k == '软件':
#         luanjian = GetMingxi(v)
#     elif k == '外包':
#         waibao = GetMingxi(v)
#     elif k == '推广费':
#         tuiguang = GetMingxi(v)
#     elif k == '工资':
#         gongzi = GetMingxi(v)
#     elif k == '库管':
#         kuguan = GetMingxi(v)
#     elif k == '办公':
#         bangong = GetMingxi(v)
#     elif k == '其他':
#         qita = GetMingxi(v)
#     else :
#         pass
# 对上面的代码进行简化
for k,v in dfs.items():
    if k != '权益':          # 不处理权益
        # print(k,v)
        v = mingxi_pinyin.get(f'{k}')
        exec(f'{v}' + '=dfs.get(k)')
    else :
        pass
    
moban = pd.read_excel(r"F:\a00nutstore\008\zww08\002电商\电商考核绩效电商销售分析\电商考核表电商销售报表-模板.xlsx")
moban.set_index('平台名称',inplace = True)
for i in column_pinyin.keys():
    mingxi = GetMingxi(eval(column_pinyin.get(i)))
    moban[f'{i}'] = moban.index.map(mingxi)
moban.rename(xiangmu_pingtai,inplace = True)  #将用友的平台名称更名为报送名称

#按各平台订单数分配平台-其他、库管费用、办公费用、其他费用
#先计算各平台订单数
# fname_dd = r"F:\a00nutstore\008\zww08\002电商\电商考核绩效电商销售分析\销售主题分析_渠道_2025-04-11_21-26-06-202501-03.xlsx"
fname_dd = easygui.fileopenbox('请打开销售主题分析_渠道文件')
# df_dd = pd.read_excel(fname_dd)
qijian = easygui.enterbox('请输入期间，"2024-01"')
def getDingdan(fname,qijian):
    df0 = pd.read_excel(fname) 
    df0 = df0[~df0['类型'].isnull()]
    df1 = df0.copy()
    df1.insert(0,'平台','')
    df1 = df1.assign(平台 = np.where(df1['类型'] == '自有店铺',df1['渠道'],df1['类型']))
    #请输入期间
    # qijian = easygui.enterbox('请输入期间，"2024-01"')
    df1.insert(0,'期间',qijian)
    df2 = df1[['平台','实发单数']]
    pivot = pd.pivot_table(df2,index = '平台',aggfunc = np.sum,fill_value = 0)
    return df1,pivot
df1,pivot = getDingdan(fname_dd,qijian)
with pd.ExcelWriter(fname_dd, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df1.to_excel(writer, sheet_name=qijian)
    pivot.to_excel(writer, sheet_name='汇总')
pivot  = pivot.assign(pingtai = pivot.index.str.split('-'))
pivot.pingtai = pivot.pingtai.str[0]
pivot.pingtai = pivot.pingtai.map(dingdan_jixiao)

#是否对运费进行调整 
isno = easygui.boolbox('是否对运费进行调整')
if isno:
    yunfei_qita = easygui.enterbox('请输入运费-其他的"余额"')
    moban.loc['其他','运费'] =  float(yunfei_qita)
else :
    pass  

#工资调整
def tiaozhengGongzi(d,s):
    s1 = s.copy()
    s1 = s1.fillna(0)
    s2 = pd.Series({'拼多多':1500,'阿里巴巴':1500,'抖音':2000})
    s3 = s1 + s2
    s3 = s3.fillna(0)
    d['工资'] = s3

tiaozhengGongzi(moban,moban.工资)

pingtai_dingdan = dict(pivot.set_index('pingtai')['实发单数'])
moban['订单数']  = moban.index.map(pingtai_dingdan)
total_dingdan = moban['订单数'].sum()
# 对库管、打包成本和包装耗材按订单数进行分配
for k  in ['kuguan','dabao','haocai']:
    v = pinyin_column.get(k)    #v = '库管'
    exec(f'{k}' + '= moban[v].sum()')
    exec('moban[v] = moban.订单数 /total_dingdan *' +f'{k}')

#对其他费用和办公费用，不能归集到各平台的按订单计算，再和能直接归集的合计
def chuliQitaBangong(total_dingdan,d,i):
    s = d[i]
    s = s.fillna(0)
    s1 = s.copy()
    feiyoung_qita = s1.get('其他')
    s1.其他 = 0
    s2 = d.订单数 / total_dingdan * feiyoung_qita
    s3 = s1 + s2
    s3 = s3.fillna(0)
    d[i] = s3

for i in ['其他','办公','运费']:
    chuliQitaBangong(total_dingdan,moban,i)    
    
moban.fillna(0,inplace = True)
moban['服务客服'] = moban['服务费'] + moban['客服费']
moban['成本小计'] = moban.loc[:,'调拨成本':'客服费'].apply(np.sum,axis = 1)
moban['毛利'] = moban['销售额'] - moban['成本小计']
moban['其他小计'] = moban.loc[:,'工资':'其他'].apply(np.sum,axis = 1)
moban['净利润'] = moban['毛利'] - moban['平台推广'] - moban['其他小计']
total = moban.sum()
moban.loc['合计'] = total
moban['毛利率'] = np.where(moban['销售额'] != 0,moban['毛利']/moban['销售额'],0)
moban['净利率'] = np.where(moban['销售额'] != 0,moban['净利润']/moban['销售额'],0)
moban.insert(0,'责任人','')
moban.责任人 = moban.index.map(pingtai_people)
moban

moban.to_excel(f'电商考核表电商销售报表{qijian}.xlsx')





