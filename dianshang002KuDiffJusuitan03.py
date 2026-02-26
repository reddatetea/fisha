'''
聚水潭003库与电商002库存差异
'''
import pandas as pd
import numpy as np
import easygui
import openpyxl
import os
import re
import TcloudCunhuoDic
import bianmabiaozhunhua06
from bianmabiaozhunhua06 import chuliNumAfterHengGang,quShu,chuliFirstHenggang,bianmaBiaozhun 
from bianmabiaozhunhua06 import huizhongXiaoshoubuDiaobojia,chuliXiaoshoubuDiaobojia


#聚水潭与大库编码异常对照表 r"F:\a00nutstore\008\zww08\产成品\聚水潭与大库编码异常对照表.xlsx"
jst_daku = {'EXL100A': 'EXL100-草莓潘达',
 'PP卡扣黄色': 'PP卡扣',
 'PP卡扣粉色': 'PP卡扣',
 'PP卡扣绿色': 'PP卡扣',
 'PP卡扣紫色': 'PP卡扣',
 'PP卡扣蓝色': 'PP卡扣',
 'PP卡扣白色': 'PP卡扣',
 'EFTZ3260NA': 'EFTZ3260Na',
 'EFJC3260NA': 'EFJC3260Na',
 'EFSZ3260NA': 'EFSZ3260Na',
 'EFZW3260NA': 'EFZW3260Na',
 'EXLA5100PA': 'EXLA5100Pa',
 'EXYA5100PA': 'EXYA5100Pa',
 'EXBB5100PA': 'EXBB5100P-A',
 'XJ408S': 'XJ408',
 '16K账夹': '16KPP账夹',
 'EXWGA5100PA': 'EXWGA5100Pa',
 'EXYB5100PA': 'EXYB5100Pa',
 'EXBA5100PA': 'EXBA5100P-A',
 'EXLB5100PA': 'EXLB5100Pa',
 'EXWGB5100PA': 'EXWGB5100PA',
 '8K160G素描纸': '8K160g素描纸',
 '4K120G素描纸': '4K素描纸',
 '4K160G素描纸': '4K160g素描纸',
 '8K120G素描纸': '8K素描纸',
 'EXLA65100NB': 'EXLA5100NB'}
lst = ['仓库编码',
       '仓库',
       '存货分类 (1级)',
       '存货分类 (2级)',
       '存货分类 (3级)',
       '存货分类 (4级)',
       '存货分类 (5级)',
       '存货编码',
       '存货',
       '存货代码',
       '数量(主单位)',
       '平均单价',
       '金额',
       '数量(辅单位)',
       '计量单位组合',
       '数量(主单位).1',
       '平均单价.1',
       '金额.1',
       '数量(辅单位).1',
       '计量单位组合.1',
       '数量(主单位).2',
       '平均单价.2',
       '金额.2',
       '数量(辅单位).2',
       '计量单位组合.2',
       '数量(主单位).3',
       '平均单价.3',
       '金额.3',
       '辅数量',
       '计量单位组合.3']
lst2 = [
    'store',
    'num',
    'class01',
    'class02',
    'class03',
    'class04',
    'class05',
    'code',
    'stock',
    'content',
    'begin_ben',
    '平均单价',
    '金额',
    'begin_jian',
    '计量单位组合',
    'ruku_ben',
    '平均单价.1',
    '金额.1',
    'ruku_jian',
    '计量单位组合.1',
    'chuku_ben',
    '平均单价.2',
    '金额.2',
    'chuku_jian',
    '计量单位组合.2',
    'end_ben',
    '平均单价.3',
    '金额.3',
    'end_jian',
    '计量单位组合.3']
lst3 = [
    'class02',
    'class05',
    'code',
    'stock',
    'content',
    'begin_ben',
    'begin_jian',
    'ruku_ben',
    'ruku_jian',
    'chuku_ben',
    'chuku_jian',
    'end_ben',
    'end_jian',
]
dic_class = {'未分类': '00',
 '测试': '1',
 '产成品': '01',
 '折扣类': '0199',
 '电商': '0106',
 '新锐': '0105',
 '复印纸': '0104',
 '锐意': '0103',
 '抄本': '0102',
 '账本': '0101',
 '电商特价': '010699',
 '定制': '010608',
 '电商抄本': '010607',
 '缝线本': '010606',
 '胶套本': '010605',
 '活页芯': '010604',
 '线环本': '010603',
 '备课本': '010602',
 '美术系列': '010601',
 '新锐订制': '010599',
 '新锐处理产品': '010598',
 '新锐防近视': '010504',
 '新锐软抄': '010503',
 '新锐胶套本': '010502',
 '新锐缝线本': '010501',
 '莱特全木浆系列': '010401',
 '锐意PP缝线活页夹': '010312',
 '锐意文稿纸': '010311',
 '锐意无线胶装': '010310',
 '锐意特价产品': '010309',
 '锐意线环': '010308',
 '锐意湖北版专利作业本': '010307',
 '锐意缝线本': '010305',
 '锐意牛卡缝线': '010304',
 '锐意软抄': '010303',
 '锐意胶套本': '010302',
 '空表': '01020814',
 '订制产品': '010214',
 '抄本处理产品': '010213',
 '莱特牛皮缝线本': '010212',
 '外贸订制产品': '010211',
 '材料纸': '010210',
 '铁钉本空': '010209',
 '小学生作业本': '010208',
 '拍纸簿': '010207',
 '皮面本空': '010206',
 '双线环本空': '010205',
 '纪念册 空': '010204',
 '空': '01021413',
 '硬抄': '010202',
 '软抄': '010201',
 '经济型无碳复写单据': '010109',
 '订制账簿系列': '010108',
 '材料纸(账本）': '010107',
 '装订配件': '010106',
 '报表': '010105',
 '凭证': '010104',
 '单据': '010103',
 '账芯': '010102',
 '账簿': '010101',
 '琪丰定制': '01060806',
 '孝感思进': '01060805',
 '领创未来': '01060804',
 '佳佳定制': '01060803',
 '优本伦定制': '01060802',
 '齐达创美定制': '01060801',
 '胶装直背本': '01060704',
 '18K胶装本': '01060703',
 'B5胶装本': '01060702',
 '胶装本': '01060701',
 '卡面缝线本': '01060603',
 '牛卡缝线本': '01060602',
 '黑卡缝线本': '0106060101',
 '16K胶套本': '01060504',
 '32K胶套本': '01060503',
 'B5胶套本': '01060502',
 'A5胶套本': '01060501',
 'A4活页芯': '01060403',
 'B5活页芯': '01060402',
 'A5活页芯': '01060401',
 'PP线环本': '01060304',
 '牛卡线环本': '01060303',
 '黑卡线环本': '01060302',
 '卡面线环本': '01060301',
 '白卡软抄': '01030304',
 '特种纸软抄': '01030303',
 '32K高白软抄': '0103030202',
 '16K高白软抄': '0103030201',
 '高白软抄': '01030302',
 '16K牛皮道林软抄': '0103030103',
 '32K道林软抄': '0103030102',
 '16K道林软抄': '0103030101',
 '道林软抄': '01030301',
 '汉正街热脉小学生本': '01020828',
 '杭州科兴小学生本': '01020827',
 '人和版小学生本': '01020826',
 '山东版小学生本': '01020825',
 '佳和版小学生本': '01020824',
 '本米版小学生本': '01020823',
 '至尚版小学生本': '01020822',
 '唐山版小学生本': '01020821',
 '天津专利防近视作业本': '01020820',
 '天津东丽津南区小字本': '01020819',
 '天津版小学生本': '01020818',
 '贵州版小学生本': '01020817',
 '重庆版小学生本': '01020816',
 '成都版小学生本': '01020815',
 '华阳版小学生本': '01020813',
 '河南版小学生本': '01020812',
 '太原版小学生本': '01020811',
 '西北版小学生本': '01020810',
 '新疆版小学生本': '01020809',
 '西安版小学生本': '01020808',
 '云南版小学生本': '01020807',
 '华北版小学生本': '01020806',
 '汉办版小学生本': '01020805',
 '长沙版小学生本': '01020804',
 '广州版小学生本': '01020803',
 '东北版小学生本': '01020802',
 '昆明版小学生本': '01020801',
 '办公硬抄': '01020203',
 '卡通硬抄': '01020202',
 '精装硬抄': '01020201',
 '精品软抄': '01020104',
 '无线胶装软抄': '01020103',
 '办公软抄': '01020102',
 '卡通软抄': '0102010102',
 '无碳复写单据': '01010303',
 '多联单据': '01010302',
 '单联单据': '01010301',
 '立信孔账芯': '01010202',
 '普通账芯': '01010201',
 '皮面账簿': '01010102',
 '纸面账簿': '01010101',
 'A4纯净牛卡缝线本': '0106060203',
 '16K牛卡缝线本': '0106060202',
 '32K牛卡缝线本': '0106060201',
 '16K黑卡缝线本': '0106060102',
 '杨曙华定制': '01060807',
 '外贸': '0107',
 '肥猫定制': '01060808',
 '取蓝缝线本': '01060604',
 '锐意专利牛皮': '010306',
 '锐意专利牛皮胶装-黄内芯': '01030601',
 '锐意专利牛皮缝线-黄内芯': '01030602',
 '16K60型专利牛皮缝线': '0103060201',
 '32K60型专利牛皮缝线': '0103060202',
 'A480型专利牛皮缝线': '0103060203',
 'A6100P软线圈本': '01030801',
 'A5100P软线圈本': '01030802',
 'B5100P软线圈本': '01030803',
 '16K100型专利牛皮缝线': '0103060204',
 'B540型专利牛皮胶装': '0103060101',
 '32K40型专利牛皮胶装': '0103060102',
 '汉办普通版': '0102080501',
 '白封面防近视内芯': '0102080502',
 '牛皮封面防近视内芯': '0102080503',
 '锐意美术系列': '010313',
 '侧翻素描本': '01031301',
 '上翻素描本': '01031302',
 '牛皮素描本': '01031303',
 '卡面速写本': '01031304',
 '素描纸': '01031305',
 '图画本': '01031306',
 '16K牛卡缝线': '01030401',
 '32K牛卡缝线': '01030402',
 'A4牛卡缝线': '01030403',
 '100K盒装线环本': '01030804',
 '锐意卡面线环': '01030805',
 '长沙刘镇宇': '01021401',
 '长沙陈德辉': '01021402',
 '长沙东麦': '01021403',
 '南宁张家胜': '01021404',
 '重庆何佳洁': '01021405',
 '成都渝泰': '01021406',
 '乌鲁木齐王建国': '01021407',
 '常州杨森': '01021408',
 '株洲于世伟': '01021409',
 '武汉热脉': '01021410',
 '南昌王雪明': '01021411',
 '其他订制': '01021412',
 '16K60型牛卡缝线': '01050101',
 '32K60型牛卡缝线': '01050102',
 'A460型牛卡缝线': '01050103',
 '16K卡面缝线': '01050104',
 '32K卡面缝线': '01050105',
 '便签草稿本': '01060101',
 '电商图画本': '01060102',
 '双线环素描本': '01060103',
 '16K纯净牛卡缝线本': '0106060204',
 '32K纯净牛卡缝线本': '0106060205',
 'A4直背本': '0106070401',
 'A5直背本': '0106070402',
 'B5直背本': '0106070403',
 '新天一定制': '01060809'}
def chuli(fname, store_num):
    df = pd.read_excel(fname, skiprows=7)
    df = df.iloc[:, 1:]
    df.columns = lst
    df.columns = lst2
    if store_num == '001库':
        df = df.loc[df.store == '001']
    elif store_num == '002电商库':
        df = df.loc[df.store == '002']
    else:
        df = df.loc[(df.store == '001') | (df.store == '002')]

    df['content'] = df['end_ben'] / df['end_jian']
    df = df[df['store'] != '制表人:']
    df = df[df['store'] != '合计：']
    df = df[df['store'].notnull()]
    df = df.iloc[:, 2:]
    df = df[lst3]
    df['bianma01'] = df['class02'].map(dic_class)
    df['bianma02'] = df['class05'].map(dic_class)
    df = df.sort_values(['bianma01', 'bianma02'])
    df = df.iloc[:, :-2]
    return df
#聚水潭库存
# fname_jst = r"F:\a00nutstore\008\zww08\002电商\聚水潭\进销存—按商品_2025-04-01_19-35-54.xlsx"
fname_jst = easygui.fileopenbox('请点选本期聚水进销存文件')
df0 = pd.read_excel(fname_jst,dtype = {'商品编码':str})
df = df0[['商品编码','期末数量']]
df = df[~df['商品编码'].isna()]

pattern = r'(.*)-(\d{1,2})$'
df1 = df.copy()
#处理最右边的-，形成bianma1
df1 = df1.assign(bianma1 = df1.商品编码.apply(lambda x:chuliNumAfterHengGang(x)[0]))
df1 = df1.assign(shuliang = df1.商品编码.apply(lambda x:chuliNumAfterHengGang(x)[1]))
#处理+,形成bianma2
df1['bianma2'] = df1['bianma1'].str.split('+')
df1['len'] = df1['bianma2'].str.len()
df1 = df1.assign(shuliang1 = np.where(df1.len == 1,df1.shuliang ,df1['shuliang']/df1['len']))
df1 = df1.explode('bianma2')
df1['数量'] =  df1['期末数量']*df1['shuliang1']


#形成存货含量字典，生成所有存货的列表
fname  = easygui.fileopenbox('请点选存货档案')
content_dic = TcloudCunhuoDic.getCunhuoConcent(fname)
cunhuo_lst = list(content_dic.keys())


def isnoInCunhuoLst(s,lst):
    if s in lst:
        return True
    else :
        return False
#如果字符在列表的某个元素中，则返回该元素
def strIsinLststr(str,lst):
    lst.sort()
    for i in lst:
        if str in i:
            return i
        else :
            continue
    return None
  

df1['bianma3'] = df1['bianma2']   #为下一步处理未匹配编码做准备
df1['bianma4'] = df1['bianma2']
df11 = df1[df1.商品编码.map(lambda x:isnoInCunhuoLst(x,cunhuo_lst))]
df12 = df1[~df1.商品编码.map(lambda x:isnoInCunhuoLst(x,cunhuo_lst))]
#直接匹配639个匹配成功，另外124个没有匹配，需要另外进一步处理
#对bianma2处理成标准编码，形成bianma3
df12['bianma3'] = df12.bianma2.map(chuliFirstHenggang)
bianma3 =  df12.bianma3.map(bianmaBiaozhun)
df12['bianma3'] = bianma3
df12['bianma4'] = df12['bianma3']



# df12.bianma3.map(content_dic)
df121 = df12[df12.bianma3.map(lambda x:isnoInCunhuoLst(x,cunhuo_lst))]
df122 = df12[~df12.bianma3.map(lambda x:isnoInCunhuoLst(x,cunhuo_lst))]
df121 = df121.assign(bianma4 = df121.bianma3)
df122 = df122.assign(bianma4 = df122.bianma3.apply(lambda x:strIsinLststr(x,cunhuo_lst)))
df122 = df122.assign(bianma4 = np.where(df122.bianma4.isin(['',None,np.nan,'NaN']),df122.bianma3.map(jst_daku),df122.bianma4))
#未匹配
df_weipipei = df122[df122.bianma4.isna()]
df122 = df122.assign(bianma4 = np.where(df122.bianma4.isin(['',None,np.nan,'NaN']),df122.bianma3,df122.bianma4))
df_jst0 = pd.concat([df11,df121,df122])
df_jst0


df_jst0.to_excel('df_jst0.xlsx',index = False)

df_jst = pd.pivot_table(df_jst0,index = 'bianma4',values = '数量',aggfunc = 'sum')
df_jst = df_jst.reset_index()
df_jst



df_jst.数量.sum()


#大库库存
# fname_daku = r"F:\a00nutstore\008\zww08\产成品\收发存汇总表2025-3-25.xlsx"
fname_daku = easygui.fileopenbox('请点选002库产成品收发存汇总表')
df_daku0 = chuli(fname_daku,'002电商库')
df_daku = df_daku0[['code','end_ben']]
pivot = pd.merge(df_jst,df_daku,how = 'outer',left_on = 'bianma4',right_on = 'code')


pivot.to_excel('pivot.xlsx',index = False)

result = pivot.copy()
result =  result.assign(code = np.where(result.code.isin(['None',np.nan,'']),result.bianma4,result.code))
result =  result.assign(bianma4 = np.where(result.bianma4.isin(['None',np.nan,'']),result.code,result.bianma4))
result.数量 = result.数量.fillna(0)
result.end_ben = result.end_ben.fillna(0)
result = result[['code','数量','end_ben']]
result = result.rename(columns = {'code':'存货编码','数量':'聚水潭','end_ben':'002库'})
result.insert(1,'差异',result['聚水潭']-result['002库'])


df_panyin = result[result['差异'] > 0]
df_pankui = result[result['差异'] < 0 ]
df_xingtong = result[result['差异'] == 0 ]
def AddTotal(d):
    total = d.sum()
    total[0] = '合计'
    d.loc['合计'] = total
    return d
df_panyin = AddTotal(df_panyin)
df_pankui = AddTotal(df_pankui)
df_xiangtong = AddTotal(df_xingtong)
    
qijian = easygui.enterbox('请输入期间')
fname_diff = f'大库与聚水潭库存差异明细-{qijian}.xlsx'
wb = openpyxl.Workbook()
wb.save(fname_diff)

with pd.ExcelWriter(fname_diff, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    
    result.to_excel(writer, sheet_name='总差异',  index=False)
    df_panyin.to_excel(writer, sheet_name='盘盈',  index=False)
    df_pankui.to_excel(writer, sheet_name='盘亏',  index=False)
    df_xiangtong.to_excel(writer, sheet_name='相同',  index=False)
    
os.startfile(fname_diff)
