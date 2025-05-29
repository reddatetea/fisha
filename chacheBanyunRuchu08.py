import re
import os
import easygui
import numpy as np
import pandas as pd
import openpyxl

YUANGONGS =  ['刘革红', '黄康', '吴长江', '李城', '胡国华', '代朝威','黄志桥','周宗华','魏道和','临时工','向明瑞','周昌正','艾金华']
dic_columns = {'fapei':'项目','date': '日期',
 'chejian': '车间',
'customer':'客户',
 'jian': '件数1',
 'gongzhong': '工种',
 'gongzhong1': '工种1',
 'people': '人员',
 'jian2': '件数2'}
gongzhongs  = ['发库叉车', '发库搬运', '配货叉车', '配货搬运','入库叉车','入库搬运']

banyun_jianchen = {'吴':'吴长江','黄':'黄志桥','代':'代朝威','李':'李城','刘':'刘革红','邹':'周宗华','临':'临时工','向':'向明瑞','周':'周昌正','艾':'艾金华','工':'车间'}
chache_jianchen = {'刘':'刘革红','黄':'黄康','胡':'胡国华','周':'周宗华','魏':'魏道和'}

dic = dict(zip(['序号', '日期', '客户', '件数','车间'],
['xuhao', 'date', 'customer', 'jian','customer']))
#入库
def chuliRuChu(fname,start_riqi,end_riqi):
    
    df = pd.read_excel(fname,sheet_name = '2024',header = 1)
    
    columns = [i  for i  in  df.columns.to_list() if i in ['序号', '日期', '车间','客户', '件数','发库叉车','发库搬运','配货叉车','配货搬运','入库叉车','入库搬运']]
    df = df[columns]
    df = df[~df['件数'].isna()]
    df = df.rename(columns = dic)
    df = df[df['date'].between(start_riqi,end_riqi) ]
    #删除空行
    # df = df.dropna()
     
    df = df.iloc[:,1:]
    column_lst = [i for i in gongzhongs if (('入库' in i ) or ('发库' in i ) or ('配货' in i)) and (i in df.columns.to_list())]
    for i in column_lst:
        df[i] = df[i].fillna('')
        df[i] = df[i].map(list)
        df = df.explode(i)
    for i in column_lst:
        if '搬运' in i:
            df[i] = df[i].map(banyun_jianchen)
        else:
            df[i] = df[i].map(chache_jianchen)
    df = df.set_index(['date','customer','jian'])
    df = df.stack()
    df = df.reset_index()
    df.columns = ['date','chejian','jian','gongzhong','people']
    df = df[df.people.isin(YUANGONGS)]
    data = []
    gp = df.groupby(['date','chejian','jian','gongzhong'])
    for k,v in gp:
        v = v.assign(jian2 = v.jian/len(v))
        data.append(v)
    yuanbiao = pd.concat(data)
    df6 = yuanbiao.assign(fapei = yuanbiao.gongzhong.str[:2],gongzhong1 = yuanbiao.gongzhong.str[2:])
    df6 = df6[[ 'fapei',
       'gongzhong1','people', 'jian2']]
    pivot = pd.pivot_table(df6,index = ['fapei','gongzhong1','people'],aggfunc = 'sum')
    pivot1 = pd.pivot_table(df6,index = 'people',columns = ['fapei','gongzhong1'],aggfunc = 'sum')
    return yuanbiao,df6,pivot1

def chuliColumnName(df,dic_columns):
    df = df.rename(columns = dic_columns)
    return df

def main():
    path = r"F:\a00nutstore\006\zw\产成品出入库工作记录"
    start_riqi = pd.Timestamp(easygui.enterbox('请输入开始日期"2024-12-26"'))
    riqi = easygui.enterbox('请输入结束日期"2025-01-25"')
    end_riqi = pd.Timestamp(riqi)
    #ruku 出库
    fname_chuku = r"F:\a00nutstore\006\zw\产成品出入库工作记录\仓库日常出库工作记录.xlsx"
    yuanbiao_chuku, df_chuku, pivot_chuku = chuliRuChu(fname_chuku,start_riqi,end_riqi)
    df_chuku = chuliColumnName(df_chuku,dic_columns)
    pivot_chuku1 = pd.pivot_table(df_chuku, index='人员', columns=['项目', '工种1'], aggfunc='sum', margins=True,
                                           margins_name='小计')

    #ruku 入库2025
    fname_ruku = r"F:\a00nutstore\006\zw\产成品出入库工作记录\仓库日常入库工作记录.xlsx"
    yuanbiao_ruku, df_ruku, pivot_ruku = chuliRuChu(fname_ruku,start_riqi,end_riqi)
    df_ruku = chuliColumnName(df_ruku,dic_columns)
    pivot_ruku1 = pd.pivot_table(df_ruku, index='人员', columns=['项目', '工种1'], aggfunc='sum', margins=True,
                                           margins_name='小计')
    
    total = pd.concat([df_ruku, df_chuku])
    total = total.rename(columns={'fapei': '项目', 'gongzhong1': '工种', 'people': '人员', 'jian2': '件数'})
    # total = total.astype({'件数2': 'int'})
    pivot = pd.pivot_table(total, index='人员', columns=['项目', '工种1'], aggfunc='sum', margins=True,
                           margins_name='小计')
    
    qijian = easygui.enterbox('请输入期间"2025-01"')
    fname_result = os.path.join(path, f'叉车搬运工时统计表-{qijian}({riqi}).xlsx')
    wb = openpyxl.Workbook()
    wb.save(fname_result)
    with pd.ExcelWriter(fname_result, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        yuanbiao_ruku.to_excel(writer, sheet_name='入库原表1', index=False)
        df_ruku.to_excel(writer, sheet_name='入库原表2', index=False)
        pivot_ruku1.to_excel(writer, sheet_name='入库汇总')
        yuanbiao_chuku.to_excel(writer, sheet_name='发库原表1', index=False)
        df_chuku.to_excel(writer, sheet_name='发库原表2', index=False)
        pivot_chuku1.to_excel(writer, sheet_name='出库汇总')
        pivot.columns= pivot.columns.droplevel(0)
        pivot.columns.delete(0)
        pivot = pivot.reset_index()
        pivot.to_excel(writer, sheet_name='汇总')
    os.startfile(fname_result)

if __name__ == "__main__":
    main()
    

    

