'''
根据多特申通或中通对账单计算电商各平台运费
每月添加到多特运费文件中
并与聚水潭里运费进行对比，找出异常
考虑拦截
可分别对申通和中通的运费进行处理
'''
import pandas as pd
import numpy as np
import easygui
import os



# lst = ['莱特纸品企业店', '莱特纸品官方旗舰店', '莱特纸品', '湖北双佳纸品有限公司']
#多特申通对账单
fname = easygui.fileopenbox('请点选多特申通或中通面单对账文件')

if '申通'  in fname :
    pingtai = '申通'
    df_ = pd.read_excel(fname,sheet_name = None)
    if '拦截退回' in df_.keys():
        df0 = pd.read_excel(fname,sheet_name = 0, dtype = {'运单号':"str"},engine='openpyxl')
        #拦截
        df1 = pd.read_excel(fname,sheet_name = 1,dtype = {'日期':'datetime64[ns]','快递单号':"str"})
        # df0['日期'] = df0['日期'].astype('datetime64[ns]')
        df1['日期'] = df1['日期'].ffill()
        df1 = df1.rename(columns = {'日期':'业务时间','快递单号':'运单号','金额':'运费'})
        df1 = df1[['业务时间','运单号','运费']]
        df = pd.concat([df0,df1])
    else :
        df = pd.read_excel(fname,sheet_name = 0, dtype = {'运单号':"str"},engine='openpyxl')
    
else :
    pingtai = '中通'
    #fname = r"F:\a00nutstore\008\zww08\电商\快递\多特运费中通面单对账单2025-02.xlsx"
    df = pd.read_excel(fname,sheet_name = 0, dtype = {'运单号':"str"},engine='openpyxl')  #,'寄件日期':"datetime64[ns]"
    df = df.rename(columns = {'寄件日期':'业务时间'})
    df = df[df['寄件客户'] == '双佳纸品']
    df['业务时间'] = pd.to_datetime(df['业务时间'])
        
path,_ = os.path.split(fname)
os.chdir(path)
df.insert(0,"期间",df['业务时间'].astype('str').str[:7])
# df.insert(0,'物流','中通')   
df.insert(0,'物流','申通')  
df = df[~df['运单号'].isin(['None',''])]
df['不含税运费'] = round(df['运费']/1.06,2)
qijian = df.iloc[-1,1]


#聚水潭主题快递分析\发货统计\按渠道\按店铺 \按日期\ 按仓库全选
# fname_jstkd  = r"F:\a00nutstore\008\zww08\电商\快递\聚水潭快递发货数量统计明细202501-202502_2025-03-07_22-05-13.xlsx"
fname_jstkd = easygui.fileopenbox('请点选聚水潭主题快递分析文件"发货统计-按渠道-按店铺=按日期-按仓库全选"')
#聚水潭回收快递单 设置\快递电子面单及打印设置\单号回收记录
# fname_jstHuishou = r"F:\a00nutstore\008\zww08\电商\快递\聚水潭导出单号回收详情20240301-20250331_2025-03-08_11-33-50.xlsx"
fname_jstHuishou = easygui.fileopenbox('请点选聚水潭导出单号回收详情文件"')
def getDanhaoPingtaiDic(fname):
    df_jstkd = pd.read_excel(fname,dtype = {'订单编号':"str",'进出仓单号':"str",'出库日期':"datetime64[ns]",'快递单号':"str"})
    qudaos = ['1688-双佳',
              '拼多多-旗舰店',
              '惊喜-双佳',
              '淘宝-企业店',
              '抖音-企业店',
              '京东-旗舰店']                   
    qudao = (
        np.where(df_jstkd['渠道'].isin(qudaos),df_jstkd['渠道'],'分销商')
               )
    df_jstkd.insert(0,'qudao',qudao)
    dic = dict(zip(df_jstkd['快递单号'],qudao))
    return df_jstkd,dic

def getHuishouDic(fname):
    df = pd.read_excel(fname,dtype = {'快递单号':"str",'内部单号':"str",'订单店铺ID':"str",'商家编码':"str",'取号店铺ID':"str",'订单店铺ID':"str"})
    df =  df[df['回收状态'] == '成功']
    dic = dict(zip(df['快递单号'],df['取号店铺名称']))
    return df,dic
    
df_justkd,dic = getDanhaoPingtaiDic(fname_jstkd)
df_huishouKd,dic_huishou  = getHuishouDic(fname_jstHuishou)   


#将聚水潭快递主题分析中的快递单号与平台制作字典
df.insert(1,'平台',df['运单号'].map(dic))
df['平台'] = df['平台'].fillna('待查')   #第一次匹配不到的作为"待查"
df['平台'] = np.where(df['平台'] == '待查',df['运单号'].map(dic_huishou),df['平台'])
df['平台'] = df['平台'].fillna('待查')   #再次匹配不到的作为"待查"
df_zhong_sheng = df.copy()
df_zhong_sheng = df_zhong_sheng.rename(columns = {'省':'省份'})
df_zhong_sheng = df_zhong_sheng[['物流','平台','期间','业务时间','运单号','省份','结算重量','运费','不含税运费']]
df_daicai = df.copy()
df_daicai = df_daicai[df_daicai['平台'] == '待查']
#按聚水潭上的平台进行汇总
dingdan = df.groupby('平台')['运单号'].count()
yunfei = df.groupby('平台')['运费'].sum()
df_yunfei = pd.concat([dingdan,yunfei],axis = 1)
df_yunfei['不含税运费']  = round(df_yunfei['运费']/1.06,2)
total = df_yunfei.sum()
df_yunfei.loc['合计'] = total

#是否将本次对账文件写入对账单
isno = easygui.boolbox('是否将本次对账文件写入对账单')
if isno:
    with pd.ExcelWriter(fname,engine='openpyxl',mode='a', if_sheet_exists='overlay')  as writer:
        df.to_excel(writer, sheet_name = f'{pingtai}平台',index = False)
        df_daicai.to_excel(writer, sheet_name = '待查',index = False)
        df_yunfei.to_excel(writer, sheet_name = '汇总')
    df_daicai.to_excel(f'{qijian}{pingtai}对账单中聚水潭没有查到的订单号明细.xlsx', sheet_name = '待查',index = False)
else :
    pass     


fname_duote_huizhong = r"F:\a00nutstore\008\zww08\电商\快递\多特运费申通中通汇总.xlsx"
isno = easygui.boolbox('是否将本期中通运费数据添加到多特运费汇总文件中')
if isno:
    df_0 = pd.read_excel(fname_duote_huizhong,sheet_name = pingtai)
    max_row = df_0.shape[0]  
    with pd.ExcelWriter(fname_duote_huizhong, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name = pingtai, startrow=max_row + 1, header=False, index=False)
     
    # #添加申通中通数据
    df_zhong_sheng0 = pd.read_excel(fname_duote_huizhong,sheet_name = '申通中通')
    max_row1 = df_zhong_sheng0.shape[0]  #申通中通
    with pd.ExcelWriter(fname_duote_huizhong, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df_zhong_sheng.to_excel(writer, sheet_name='申通中通', startrow=max_row1 + 1, header=False, index=False)
else :
    pass
os.startfile(fname_duote_huizhong)                        
   

        
 


