'''
对叉车的搬运的工时进行统计
'''
import re
import os
import easygui
import numpy as np
import pandas as pd
import openpyxl

YUANGONGS =  ['刘革红', '黄康', '吴长江', '李城', '胡国华', '代朝威','黄志桥']
dic_columns = {'fapei':'项目','date': '日期',
 'chejian': '车间',
 'jian': '件数1',
 'gongzhong': '工种',
 'gongzhong1': '工种1',
 'people': '人员',
 'jian2': '件数2'}

def chuliRuChu(ruchu,fname,start_riqi,end_riqi):
    # fname = r"F:\a00nutstore\006\zw\产成品出入库工作记录\仓库日常入库工作记录.xlsx"
    df = pd.read_excel(fname,header = 1)
    df = df.rename(columns = {'货物数量（件）':'jian','货物件数':'jian','日期':'date','车间':'address','单位名称':'address'})
    df = df[[i  for i  in  df.columns.to_list() if (i != '货物数量(托)') and  (i != '序号') and ('Unnamed' not in i)]]
    #删除空行
    df = df[df['jian'].notna()]
    df = df[df['date'].between(start_riqi,end_riqi) ]
    df = df.set_index(['date','address','jian'])
    df = df.stack()
    df = df.reset_index()
    df.columns = ['date','chejian','jian','gongzhong','people']
    df = df[df.people.isin(YUANGONGS)]
    df.gongzhong = df.gongzhong.str.replace(r'\d+','',regex = True)
    data = []
    gp = df.groupby(['date','chejian','jian','gongzhong'])
    for k,v in gp:
        v = v.assign(jian2 = v.jian/len(v))
        data.append(v)
    yuanbiao = pd.concat(data)
    #原表
    # df5.to_excel(f'叉车搬运工时加工后{ruku}原表01.xlsx',index = False)
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
    start_riqi = pd.Timestamp(easygui.enterbox('请输入入库开始日期：格式为"2024-9-26"'))
    end_riqi = pd.Timestamp(easygui.enterbox('请输入入库结束日期：格式为"2024-10-25"'))
    path = r"F:\a00nutstore\006\zw\产成品出入库工作记录"
    os.chdir(path)
    lst = os.listdir(path)
    for file in lst:
        if file == '仓库日常入库工作记录.xlsx':
            ruchu = '入库'
            fname = os.path.join(path, file)
            yuanbiao_ruku, df_ruku, pivot_ruku = chuliRuChu(ruchu, fname, start_riqi, end_riqi)
            yuanbiao_ruku = chuliColumnName(yuanbiao_ruku, dic_columns)
            df_ruku = chuliColumnName(df_ruku, dic_columns)
            pivot_ruku1 = pd.pivot_table(df_ruku, index='人员', columns=['项目', '工种1'], aggfunc='sum', margins=True,
                                         margins_name='小计')

        elif file == '仓库日常出库工作记录.xlsx':
            ruchu = '出库'
            fname = os.path.join(path, file)
            yuanbiao_chuku, df_chuku, pivot_chuku = chuliRuChu(ruchu, fname, start_riqi, end_riqi)
            yuanbiao_chuku = chuliColumnName(yuanbiao_chuku, dic_columns)
            yuanbiao_chuku = yuanbiao_chuku.rename(columns={'车间': '客户'})
            df_chuku = chuliColumnName(df_chuku, dic_columns)
            df_chuku = df_chuku.rename(columns={'车间': '客户'})
            pivot_chuku1 = pd.pivot_table(df_chuku, index='人员', columns=['项目', '工种1'], aggfunc='sum', margins=True,
                                          margins_name='小计')
        else:
            continue

    total = pd.concat([df_ruku, df_chuku])
    total = total.rename(columns={'fapei': '项目', 'gongzhong1': '工种', 'people': '人员', 'jian2': '件数'})
    # total = total.astype({'件数2': 'int'})
    pivot = pd.pivot_table(total, index='人员', columns=['项目', '工种1'], aggfunc='sum', margins=True,
                           margins_name='小计')

    fname_result = os.path.join(path, '叉车搬运工时统计表.xlsx')
    wb = openpyxl.Workbook()
    wb.save(fname_result)
    with pd.ExcelWriter(fname_result, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        yuanbiao_ruku.to_excel(writer, sheet_name='入库原表1', index=False)
        df_ruku.to_excel(writer, sheet_name='入库原表2', index=False)
        pivot_ruku1.to_excel(writer, sheet_name='入库汇总')
        yuanbiao_chuku.to_excel(writer, sheet_name='出库原表1', index=False)
        df_chuku.to_excel(writer, sheet_name='出库原表2', index=False)
        pivot_chuku1.to_excel(writer, sheet_name='出库汇总')
        pivot.to_excel(writer, sheet_name='汇总')

    os.startfile(fname_result)

if __name__ == '__main__':
    main()







