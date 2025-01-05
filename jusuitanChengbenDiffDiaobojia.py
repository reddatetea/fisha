'''
聚水潭成本价与销售部内部调拨价的差异
'''


import pandas as pd
import numpy as np
import easygui
import openpyxl
import os




lst0_jst= ['图片',
 '款式编码',
 '商品编码',
 '商品名称',
 '商品简称',
 '颜色及规格',
 '颜色',
 '规格',
 '基本售价',
 '成本价',
 '采购价',
 '市场|吊牌价',
 '品牌',
 '分类',
 '虚拟分类',
 '商品标签',
 '国标码',
 '供应商名称',
 '重量',
 '长',
 '宽',
 '高',
 '体积',
 '单位',
 '商品状态',
 '库存同步',
 '备注',
 '库容下限',
 '库容上限',
 '溢出数量',
 '标准装箱数量',
 '标准装箱体积',
 '主仓位',
 '其它价格1',
 '其它价格2',
 '其它价格3',
 '其它属性1',
 '其它属性2',
 '其它属性3',
 '修改时间',
 '创建时间',
 '创建人']
lst1_jst = [ '款式编码',
 '商品编码',
 '商品名称',
  '基本售价',
 '成本价',
  '重量',
 '其它属性1',
 '其它属性2',
]
lst0_xsb = ['Unnamed: 0',
 '类别',
 '存货编码',
 '存货名称',
 '规格型号',
 '计价方式',
 '所属类别',
 '计量单位',
 '存货代码',
 '最新成本',
 '建档日期',
 '修改日期',
 '汉正街批发价',
 '汉口北批发价',
 '外地批发价',
 '内部调拨价']
lst1_xsb = [ '类别',
 '存货编码',
 '存货名称',
 '规格型号',
  '所属类别',
 '计量单位',
 '内部调拨价']



#聚水潭成本价格
# fname_juusitan = r"F:\a00nutstore\008\zw08\电商\聚水潭\聚水潭商品资料成本2024-12-7.xlsx"
fname_juusitan = easygui.fileopenbox('请点选"聚水潭商品资料成本"文件')
df_jusuitan0 = pd.read_excel(fname_juusitan,dtype ={'商品编码':str})
df_jusuitan1 = df_jusuitan0[lst1_jst]
path,fname = os.path.split(fname_juusitan)
os.chdir(path)
#销售部内部调拨价格汇总
# fname_xiaoshou = r"F:\a00nutstore\008\zw08\销售部价格\内部调拨价格表2024.10.19_汇总.xlsx"
fname_xiaoshou = easygui.fileopenbox('请点选"销售部内部调拨价"文件')
df_xiaoshou0 = pd.read_excel(fname_xiaoshou,dtype ={'存货编码':str})
df_xiaoshou1 = df_xiaoshou0[lst1_xsb]



#合并上面两表
diff = pd.merge(df_jusuitan1,df_xiaoshou1,how = 'left',left_on = '商品编码',right_on = '存货编码')
diff['差额'] = diff['内部调拨价'] - diff['成本价']
diff1= diff[diff['差额'] != 0]
diff1 = diff1[['商品编码','商品名称','成本价','内部调拨价','差额']]



wb=openpyxl.Workbook()
fname_result = r"聚水潭成本与内部调拨价差额.xlsx"
wb.save(fname_result)

with pd.ExcelWriter(fname_result, engine='openpyxl',mode='a', if_sheet_exists='overlay')  as writer:
    df_jusuitan0.to_excel(writer, sheet_name = '聚水潭',index = False)
    df_xiaoshou0.to_excel(writer, sheet_name = '销售',index = False)
    diff.to_excel(writer, sheet_name = '差异原表',index = False)
    diff1.to_excel(writer, sheet_name = '差异',index = False)
os.startfile(fname_result)
    

