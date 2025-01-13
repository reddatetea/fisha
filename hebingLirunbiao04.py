'''
根据五家新公司、双佳003、莱特及大凭证上的成本，自动生成合并利润表当月数
根据选取的上月数，计算出合并利润表的累计数
2025-1-11尝试将各公司的费用明细写入到多栏式费用明细账中
'''
import pandas as pd
import os
import easygui
import re
import openpyxl


#五家新公司利润表统一格式
dic = {'一、产品销售收入': '一、营业收入',
 '减：销售成本': '减：营业成本',
 '销售税金及附加': '营业税金及附加',
 '减：销售费用': '减：销售费用',
 '管理费用': '管理费用',
 '财务费用': '财务费用',
 '营业外收入': '加：营业外收入',
 '其它': '投资收益（损失以"-"填列）',
 '减：营业外支出': '减：营业外支出'}

def getShouruChenben(fname):
    df = pd.read_excel(fname,header = 3)
    df.columns = ['xiangmu','benyue','leiji']
    df = df.fillna(0)
    df['xiangmu']= df['xiangmu'].str.strip()
    df = df.set_index('xiangmu')
    shouru = df.loc['一、营业收入','benyue']
    chenben = df.loc['减：营业成本','benyue']    
    guanlifeiyong = df.loc['管理费用','benyue']
    xiaoshoufeiyong = df.loc['销售费用','benyue']
    yingyesuijin = df.loc['营业税金及附加','benyue']
    caiwufeiyong = df.loc['财务费用','benyue']
    feiyong = [guanlifeiyong,xiaoshoufeiyong,yingyesuijin,caiwufeiyong]
    return df,shouru,chenben,feiyong

def getLaite(fname_laite):
    df_laite = pd.read_csv(fname_laite,encoding = 'GBK')  #dtype = {'C本 月 数':'float64','D本年累计数':'float64'}
    df_laite = df_laite.fillna(0)
    df_laite.drop('B行次',inplace = True,axis = 1)
    df_laite.columns = ['xiangmu','benyue','leiji']
    df_laite.xiangmu = df_laite.xiangmu.str.strip()
    df_laite = df_laite.set_index('xiangmu')
    df_laite.benyue = df_laite.benyue.str.replace(',','')
    df_laite.benyue = df_laite.benyue.astype('float64')
    df_laite.leiji = df_laite.leiji.str.replace(',','')
    df_laite.leiji = df_laite.leiji.astype('float64')
    df_laite.benyue = df_laite.benyue.fillna(0)
    df_laite.leiji = df_laite.leiji.fillna(0)
    shouru = df_laite.loc['一，产品销售收入','benyue']
    qitalirun = df_laite.loc['加：其他业务利润','benyue']
    df_laite.loc['一，产品销售收入','benyue'] = shouru + qitalirun
    df_laite.index = df_laite.index.map({'一，产品销售收入': '一、营业收入',
     '减：产品销售成本': '减：营业成本',
     '产品销售税金及附加': '营业税金及附加',
     '产品销售费用': '销售费用',
     '减：管理费用': '管理费用',
     '财务费用': '财务费用',
     '加：投资收益': '投资收益（损失以"-"填列）',
     '三，营业利润': '二、营业利润（亏损以"－"号填列）',
     '营业外收入': '加：营业外收入',
     '减：营业外支出': '减：营业外支出',
     '四，利润总额': '三、利润总额（亏损总额以"－"号填列）',
     '减：所得税': '减：所得税费用',
     '五，净利润': '四、净利润（净亏损以"－"号填列）'})
    
    df_laite = df_laite.loc[df_laite.index.dropna()]
    df_laite =  df_laite.loc[ ['一、营业收入',
     '减：营业成本',
     '营业税金及附加',
     '销售费用',
     '管理费用',
     '财务费用',
     '投资收益（损失以"-"填列）',
     '二、营业利润（亏损以"－"号填列）',
     '加：营业外收入',
     '减：营业外支出',
     '三、利润总额（亏损总额以"－"号填列）',
     '减：所得税费用',
     '四、净利润（净亏损以"－"号填列）']]
    df_laite= df_laite.reindex(index =['一、营业收入',
     '减：营业成本',
     '营业税金及附加',
     '销售费用',
     '管理费用',
     '财务费用',
     '资产减值损失',
     '加：公允价值变动收益（损失以"-"填列）',
     '投资收益（损失以"-"填列）',
     '其中：对联营企业和合营企业的投资收益',
     '二、营业利润（亏损以"－"号填列）',
     '加：营业外收入',
     '减：营业外支出',
     '其中：非流动资产处置损失',
     '三、利润总额（亏损总额以"－"号填列）',
     '减：所得税费用',
     '四、净利润（净亏损以"－"号填列）',
     '五、每股收益：',
     '（一）基本每股收益',
     '（二）稀释每股收益']
    ,fill_value = 0)
    shouru = df_laite.loc['一、营业收入','benyue']
    chenben = df_laite.loc['减：营业成本','benyue']
    guanlifeiyong = df_laite.loc['管理费用','benyue']
    xiaoshoufeiyong = df_laite.loc['销售费用','benyue']
    yingyesuijin = df_laite.loc['营业税金及附加','benyue']
    caiwufeiyong = df_laite.loc['财务费用','benyue']
    feiyong = [guanlifeiyong,xiaoshoufeiyong,yingyesuijin,caiwufeiyong]
    return df_laite,shouru,chenben,feiyong

# path = r'F:\a00nutstore\008\zw08\新公司\7月财务报表'
path = easygui.diropenbox('请点选各公司财务报表所在文件夹',title = '莱特纸品的报表请先准备好csv格式的文件')
_,fname = os.path.split(path)
os.chdir(path)
pattern = r'(?P<qijian>\d+月).+'
mat = re.search(pattern,fname)

if mat:
    qijian = mat.group(1)
else :
    qijian = ''
sheet_name =  qijian + '费用明细'
filename = qijian + '利润表汇总.xlsx'
lirunLst = [i for i in os.listdir(path) if ('利润' in i) and (not i.startswith('~$')) ]

# for i in ['双佳','莱特','销售','莱新','佳广','荣佳','佳科']:
#销售收入抵减
jian_shouru = []
#销售成本抵减
jian_chengben = []
data = []
gongsi_feiyong = {}
gongsis = []
for file in lirunLst:
    if ('莱特' in file) and (file.endswith('.csv')):
        # fname_laite = r"F:\a00nutstore\008\zw08\新公司\7月财务报表\莱特202407利润表.csv"
        gongsi = 'laite_gongsi'
        df,shouru,chengben,feiyong= getLaite(file)
        data.append(df)
        jian_shouru.append(chengben)
        jian_chengben.append(chengben)
        gongsi_feiyong[gongsi] = feiyong
        gongsis.append(gongsi)
    else :
        
        if  ('双佳' in file) and ('销售' not in file):
            gongsi = 'shuangjia_gongsi'
            df,shouru,chengben,feiyong = getShouruChenben(os.path.join(path,file))
            data.append(df)
            gongsi_feiyong[gongsi] = feiyong
            gongsis.append(gongsi)
            
            # jian_shouru.append(chengben)
            # jian_chengben.append(chengben)
            
        elif  '销售' in file:
            gongsi = 'xiaoshou_gongsi'
            df,shouru,chengben,feiyong = getShouruChenben(os.path.join(path,file))
            data.append(df)
            jian_shouru.append(chengben)
            jian_chengben.append(chengben)
            gongsi_feiyong[gongsi] = feiyong
            gongsis.append(gongsi)
        elif '莱新' in file:
            gongsi = 'laixin_gongsi'
            df,shouru,chengben,feiyong = getShouruChenben(os.path.join(path,file))
            data.append(df)
            jian_shouru.append(chengben)
            jian_chengben.append(chengben)
            gongsi_feiyong[gongsi] = feiyong
            gongsis.append(gongsi)
        elif '佳广' in file:
            gongsi = 'jiaguang_gongsi'
            df,shouru,chengben,feiyong = getShouruChenben(os.path.join(path,file))
            data.append(df)
            jian_shouru.append(chengben)
            jian_chengben.append(chengben)
            gongsi_feiyong[gongsi] = feiyong
            gongsis.append(gongsi)
        elif '荣佳' in file:
            gongsi = 'rongjia_gongsi'
            df,shouru,chengben,feiyong = getShouruChenben(os.path.join(path,file))
            data.append(df)
            jian_shouru.append(chengben)
            jian_chengben.append(chengben)
            gongsi_feiyong[gongsi] = feiyong
            gongsis.append(gongsi)
        elif '佳科' in file:
            gongsi = 'jiake_gongsi'
            df,shouru,chengben,feiyong = getShouruChenben(os.path.join(path,file))
            data.append(df)
            jian_shouru.append(chengben)
            jian_chengben.append(chengben)
            gongsi_feiyong[gongsi] = feiyong
            gongsis.append(gongsi)
        else :
            continue
result0 = pd.concat(data)
benyue = result0.groupby('xiangmu',sort = False)['benyue'].sum()
leiji = result0.groupby('xiangmu',sort = False)['leiji'].sum()
result = pd.DataFrame({'benyue':benyue,'leiji':leiji},columns = ['benyue','leiji'])
result.loc['一、营业收入','benyue'] = result.loc['一、营业收入','benyue'] - sum(jian_shouru)
result.loc['减：营业成本'] = result.loc['减：营业成本','benyue'] -  sum(jian_chengben)
result.loc['二、营业利润（亏损以"－"号填列）','benyue'] = result.loc['二、营业利润（亏损以"－"号填列）','benyue'] - sum(jian_shouru) + sum(jian_chengben)
result.loc['二、营业利润（亏损以"－"号填列）','leiji'] = result.loc['二、营业利润（亏损以"－"号填列）','leiji'] - sum(jian_shouru) + sum(jian_chengben)
result = result.reset_index()
result.columns = ['项目','本月','累计']

feiyongmingxis = []
for gongsi in gongsis:
    feiyong = gongsi_feiyong.get(gongsi)
    guanlifeiyong = feiyong[0]
    xiaoshoufeiyong = feiyong[1]
    yingyesuijin = feiyong[2]
    caiwufeiyong = feiyong[3]
    gongsifeiyong = [gongsi,guanlifeiyong,xiaoshoufeiyong,yingyesuijin,caiwufeiyong]
    feiyongmingxis.append(gongsifeiyong)
df_feiyong = pd.DataFrame(feiyongmingxis,columns = ['公司','管理费用','销售费用','营业税金及附加','财务费用'])
df_feiyong['公司'] = df_feiyong['公司'].map({'jiaguang_gongsi': '佳广',
 'rongjia_gongsi': '荣佳',
 'shuangjia_gongsi': '双佳',
 'jiake_gongsi': '佳科',
 'laixin_gongsi': '莱新',
 'xiaoshou_gongsi': '销售',
 'laite_gongsi': '莱特'})
# 指定排序的顺序
order = ['双佳', '莱特', '销售', '莱新', '佳广', '荣佳', '佳科']

# 将'Name'列转换为分类类型，并指定排序顺序
df_feiyong['公司'] = pd.Categorical(df_feiyong['公司'], categories=order, ordered=True)
df_feiyong = df_feiyong.sort_values(by='公司')
df_feiyong['合计'] = df_feiyong['管理费用'] + df_feiyong['销售费用'] + df_feiyong['营业税金及附加'] + df_feiyong['财务费用'] 
df_feiyong['公司'] = df_feiyong['公司'].astype('str')
total = df_feiyong.sum().to_frame().T
df_feiyong1 = pd.concat([df_feiyong,total])
df_feiyong1.iloc[-1,0] = '合计'

wb=openpyxl.Workbook()
wb.save(filename)
with pd.ExcelWriter(filename,engine='openpyxl',mode='a', if_sheet_exists='overlay')  as writer:
    df_feiyong1.to_excel(writer, sheet_name = sheet_name,index = False)
    result.to_excel(writer, sheet_name = '汇总',index = False)

wb=openpyxl.load_workbook(filename)
ws = wb['Sheet']
wb.remove_sheet(ws)
ws = wb['汇总']
wb.save(filename)
os.startfile(filename)

