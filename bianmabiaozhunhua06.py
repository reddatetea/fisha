'''
聚水潭编码和大库编码标准化
'''
import pandas as pd
import numpy as np
import easygui
import openpyxl
import os
import re
import jusuitanDiffChuku04

pattern = r'(.*)-(\d{1,2})$'
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

#聚水潭上编码与销售部的编码对照表
jusuitanBianbaDuizhaobiao = {'16K账夹': '16KPP账夹',
 '4K120G素描纸': '4K素描纸',
 '4K160G素描纸': '4K160g素描纸',
 '8K120G素描纸': '8K素描纸',
 '8K160G素描纸': '8K160g素描纸',
 'ATHA460/THA460': 'ATHA460N',
 'B5-152页书芯': 'B5-152页书芯t',
 'B5-68页书芯': 'B5-68页书芯t',
 'EFB3260HB': 'EFB3260HB（黑卡）a',
 'EFB3260NB': 'EFB3260NB（电商）',
 'EFL3260HB': 'EFL3260HB（黑卡）a',
 'EFL3260NB': 'EFL3260NB（电商）',
 'EFWG3260NB': 'EFWG3260NB（电商）',
 'EFY3260NB': 'EFY3260NB（电商）',
 'EFZY3260-C': 'EFZY3260t-C',
 'EJDCB540': 'EJDCB540t',
 'EJDWB550-拼音': 'EJDWB550-拼音t',
 'EJDWB550-英高': 'EJDWB550-英高t',
 'EJL16150': 'EJL16150-B01',
 'EJL1680': 'EJL1680-C01',
 'EJRJ1680': 'EJRJ1680-C01',
 'EJRJ3280': 'EJRJ3280-B01',
 'EJY16150': 'EJY16150-A01',
 'EJY1680': 'EJY1680-A01',
 'EJYBB520-红米字': 'EJYBB520-红米字t',
 'EJYBB520-绿米字': 'EJYBB520-绿米字t',
 'EJYBB520-绿田字': 'EJYBB520-绿田字t',
 'EJZW16150': 'EJZW16150-B01',
 'EJZW1680': 'EJZW1680-A01',
 'EXA4PP-白色': 'EXA4PP',
 'EXA4PP-粉色': 'EXA4PP',
 'EXA4PP-黄色': 'EXA4PP',
 'EXA4PP-蓝色': 'EXA4PP',
 'EXA4PP-绿色': 'EXA4PP',
 'EXA4PP-紫色': 'EXA4PP',
 'EXA5PP-白色': 'EXA5PP',
 'EXA5PP-粉色': 'EXA5PP',
 'EXA5PP-黄色': 'EXA5PP',
 'EXA5PP-蓝色': 'EXA5PP',
 'EXA5PP-绿色': 'EXA5PP',
 'EXA5PP-紫色': 'EXA5PP',
 'EXB5PP-白色': 'EXB5PP',
 'EXB5PP-粉色': 'EXB5PP',
 'EXB5PP-黄色': 'EXB5PP',
 'EXB5PP-蓝色': 'EXB5PP',
 'EXB5PP-绿色': 'EXB5PP',
 'EXB5PP-紫色': 'EXB5PP',
 'EXLA580HB-C': 'EXLA580HBt-C',
 'EXLB580HB-C': 'EXLB580HBt-C',
 'EXLB580NB-C': 'EXLB580NBt-C',
 'LCJB64200NB/LCJB64200HB': 'LCJB64200NB',
 'LCJB64250NB/LCJB64250HB': 'LCJB64250NB',
 'N0516': 'N0516办',
 'N0780': 'N0780（电商）a',
 'N0781（单线）': 'N0781单线电商t',
 'N0781（方格）': 'N0781方格电商t',
 'N0782（单线）': 'N0782单线t',
 'N0782（双竖线）': 'N0782双竖线电商t',
 'N0783': 'N0783电商t',
 'N0784（300格）': 'N0784（300格）电商a',
 'N0784（400格）': 'N0784（400格）电商a',
 'N0785': 'N0785电商t',
 'N0786': 'N0786（电商）a',
 'N0787': 'N0787（电商）a',
 'N0791': 'N0791（电商）',
 'N0792': 'N0792（电商）',
 'N0793': 'N0793（电商）',
 'N0794': 'N0794（电商）',
 'N0795': 'N0795（电商）',
 'N0796': 'N0796（电商）',
 'N0797': 'N0797（电商）',
 'NFC1660NH(方格）': 'NFC1660NH（方格）',
 'NFSX1660NH（单行）': 'NFSX1660NH(单行）',
 'NFSX1660NH(双竖线）': 'NFSX1660NH（双竖线）',
 'NFY1660NH(双色）': 'NFY1660NH（双色）',
 'NFZW1660NH(300格）': 'NFZW1660NH（300格）',
 'NFZW1660NH(400格）': 'NFZW1660NH（400格）',
 'NO342': 'N0342',
 'RFC-1660N（方格）': 'RFC-1660N',
 'RFL-3260': 'RFL-3260N',
 'RFYB-1660N(田字格）': 'RFYB-1660N（田字格）',
 'RFZW-1660': 'RFZW-1660N（300格）',
 'RJL-16100': 'RJL-16100-H01',
 'RJL-16200': 'RJL-16200-H02',
 'RJL-1680': 'RJL-1680-H01',
 'RJL-32150': 'RJL-32150-H01',
 'RJL-3280': 'RJL-3280-I01',
 'RJRJ-1680': 'RJRJ-1680-G01',
 'RJRJ-3280': 'RJRJ-3280-H03',
 'RJY-16150': 'RJY-16150N',
 'RJY-3280': 'RJY-3280-G01',
 'WL-1680': 'WL1680',
 'XJ408S': 'XJ408',
 '账钉（包）': '账钉'}
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
#聚水潭上的商品编码，但销售部没有编码
lst_jusuitan = ['25K账夹',
 '2006',
 '3604',
 '407',
 '504',
 '7004',
 '7006',
 '7014',
 '7016',
 '7401',
 '7402',
 '7503',
 '7601',
 '7602',
 '7605',
 '8101',
 '8102',
 'A4图画本',
 'A5-52页书芯',
 'A570克',
 'A5-92页书芯',
 'A5胶套',
 'B5-52页书芯',
 'B5-92页书芯',
 'B5胶套',
 'B5图画本',
 'BTHA460/EWTHA430',
 'BZTH-30',
 'BZTH-50',
 'CG507(彩）',
 'CGB1620',
 'CTHA460/THWA460',
 'DL1640',
 'DL1660',
 'DL1680',
 'DY-16100',
 'DY-1660',
 'DY-1680',
 'DZW-1640',
 'EFB1660HB',
 'EFL1660',
 'EFL1660-优本伦订制',
 'EFLA560',
 'EFLA560K-甜系/盐系定制',
 'EFLB560',
 'EFWG1660HB',
 'EFWG3260HB',
 'EFWGA460HB',
 'EXLA580',
 'EXLA580NB-C',
 'EXLA580-佳佳订制',
 'EXLB580',
 'EXLB580-佳佳订制',
 'H1201',
 'H1203',
 'JT-B5',
 'N0312',
 'N0315',
 'N0322',
 'N0575',
 'N0586',
 'N0587',
 'N0588',
 'N0589',
 'N0739',
 'N0740',
 'N0784（500格）',
 'N0785（纯净版）',
 'N0788',
 'N0789',
 'N0795（纯净版）',
 'N0798',
 'NFC-A460',
 'NFTZ-1660',
 'NFZW-1660（320格）',
 'PP卡扣',
 'RC-1680',
 'RFC-1660N(单线）',
 'RFC-1660N(双线）',
 'RFDL-1660N',
 'RFFG-1660N',
 'RFHX-1660N',
 'RFJC-1660',
 'RFKW-1660N',
 'RFL-16100',
 'RFL-1660',
 'RFL-1660N(纯净版）',
 'RFL-32100',
 'RFLS-1660N',
 'RFRJ-3260',
 'RFSW-1660N',
 'RFWL-1660N',
 'RFY-1660',
 'RFZW-1660N（500格）',
 'RFZZ-1660N',
 'RJL-32100',
 'RJY-1680',
 'RJZW-1680',
 'RL-16150',
 'RL-1660',
 'RL-1680',
 'RL-1680NZ',
 'RL-3280',
 'RY-1660',
 'RY-1680',
 'RY-1680NZ',
 'RZW-1680',
 '合同',
 '介绍信',
 '启用表（张）']

#聚水潭上有编码，而销售部没有，聚水潭上这部分产品的成本价
dic_jusuitan = {'16K账夹': 3.02,
 '4K120G素描纸': 4.15,
 '4K160G素描纸': 5.99,
 '8K120G素描纸': 2.05,
 '8K160G素描纸': 3.09,
 'ATHA460/THA460': 1.72,
 'B5-152页书芯': 3.34,
 'B5-68页书芯': 1.64,
 'EFB3260HB': 0.55,
 'EFB3260NB': 0.55,
 'EFL3260HB': 0.55,
 'EFL3260NB': 0.55,
 'EFWG3260NB': 0.55,
 'EFY3260NB': 0.55,
 'EFZY3260-C': 0.55,
 'EJDCB540': 0.92,
 'EJDWB550-拼音': 0.97,
 'EJDWB550-英高': 0.97,
 'EJL16150': 3.1,
 'EJL1680': 2.09,
 'EJRJ1680': 2.09,
 'EJRJ3280': 1.36,
 'EJY16150': 3.1,
 'EJY1680': 2.09,
 'EJYBB520-红米字': 0.36,
 'EJYBB520-绿米字': 0.36,
 'EJYBB520-绿田字': 0.36,
 'EJZW16150': 3.1,
 'EJZW1680': 2.09,
 'EXA4PP-白色': 2.01,
 'EXA4PP-粉色': 2.01,
 'EXA4PP-黄色': 2.01,
 'EXA4PP-蓝色': 2.01,
 'EXA4PP-绿色': 2.01,
 'EXA4PP-紫色': 2.01,
 'EXA5PP-白色': 1.23,
 'EXA5PP-粉色': 1.23,
 'EXA5PP-黄色': 1.23,
 'EXA5PP-蓝色': 1.23,
 'EXA5PP-绿色': 1.23,
 'EXA5PP-紫色': 1.23,
 'EXB5PP-白色': 1.66,
 'EXB5PP-粉色': 1.66,
 'EXB5PP-黄色': 1.66,
 'EXB5PP-蓝色': 1.66,
 'EXB5PP-绿色': 1.66,
 'EXB5PP-紫色': 1.66,
 'EXLA580HB-C': 1.1,
 'EXLB580HB-C': 1.65,
 'EXLB580NB-C': 1.65,
 'LCJB64200NB/LCJB64200HB': 1.3,
 'LCJB64250NB/LCJB64250HB': 1.6,
 'N0516': 1.02,
 'N0780': 0.8,
 'N0781（单线）': 0.8,
 'N0781（方格）': 0.8,
 'N0782（单线）': 0.8,
 'N0782（双竖线）': 0.8,
 'N0783': 0.8,
 'N0784（300格）': 0.8,
 'N0784（400格）': 0.8,
 'N0785': 0.8,
 'N0786': 0.8,
 'N0787': 0.8,
 'N0791': 0.55,
 'N0792': 0.55,
 'N0793': 0.55,
 'N0794': 0.55,
 'N0795': 0.55,
 'N0796': 0.55,
 'N0797': 0.55,
 'NFC1660NH(方格）': 1.33,
 'NFSX1660NH（单行）': 1.33,
 'NFSX1660NH(双竖线）': 1.33,
 'NFY1660NH(双色）': 1.33,
 'NFZW1660NH(300格）': 1.33,
 'NFZW1660NH(400格）': 1.33,
 'NO342': 1.07,
 'RFC-1660N（方格）': 1.26,
 'RFL-3260': 0.77,
 'RFYB-1660N(田字格）': 1.26,
 'RFZW-1660': 1.28,
 'RJL-16100': 3.64,
 'RJL-16200': 5.47,
 'RJL-1680': 2.99,
 'RJL-32150': 2.74,
 'RJL-3280': 1.8,
 'RJRJ-1680': 2.99,
 'RJRJ-3280': 1.8,
 'RJY-16150': 4.56,
 'RJY-3280': 1.8,
 'WL-1680': 0,
 'XJ408S': 0.81,
 '账钉（包）': 29.53}
def huizhongXiaoshoubuDiaobojia(isno):
    if isno:
        fname = easygui.fileopenbox('请点选"销售部内部调拨价"文件')
        path,hou= os.path.splitext(fname)
        newPath = path + '_汇总'
        fname_price = newPath + hou
        dfs = pd.read_excel(fname,sheet_name = None,dtype ={'存货编码':str})
        data = []
        for k,v in dfs.items():
            v.insert(0,'类别',k)
            data.append(v)
        df_price = pd.concat(data)
        df_xiaoshou1 = df_price[lst1_xsb]
        df_xiaoshou1.to_excel(fname_price,index = False)
    else :
        # fname = easygui.fileopenbox('请点选"销售部内部调拨价汇总"文件') 
        fname = r"F:\a00nutstore\008\zww08\销售部价格\用友软件存货编码内部调拨价12.19_汇总.xlsx"
        df_xiaoshou1 = pd.read_excel(fname,dtype = {'存货编码':str})
    return df_xiaoshou1

def quShu(str):
    regex = re.compile(r'^(\d+)\D+') 
    if str ==0 :
        return str
    else :
        try:
            str = float(str)
            return str
        except :
            mat = re.search(regex,str)
            if mat:
                str = float(mat.group(1))
                return str
            else:
                str = 0
                return str            

#将编码调整为标准
#去掉从左往右数第一个'-',如果左边的长度小于或等于4
def chuliFirstHenggang(s):
    first = s.find('-')
    if first == -1 :           #没找到'-'
        s1 = s
    else :
        if first <= 4 :
            s1 = s[:first] + s[first+1:]           
        else :
            s1 = s
    return s1

#去掉-C
def chuliGangC(s):
    if s.endswith('-C'):
        s1 = s[:-2] 
    else :
        s1 = s
    return s1
#最后一个字符是小写的a,改为大写
def chuliXiaoxieA(s):
    if s.endswith('a'):
        s1 = s[:-1] + 'A'
    else :
        s1 = s
    return s1
def chuliXiaoxieT(s):
    if s.endswith('t'):
        s1 = s[:-1] + '-T'
    else :
        s1 = s
    return s1

#替换"单行"、“双色”，“双行”、“方格”、“300格”，‘双竖线’为空
def chuliTedingZhifu(s):
    s1 = s.replace('单行','').replace('双色','').replace('双行','').replace('方格','').replace('300格','').replace('双竖线','')
    return s1
  
    
    
#将'(',")"替换为'-'
def chuliHuohao(s):
    s1 = s.replace("（",'-').replace("）",'-').replace("(",'-').replace(")",'-')
    s1 = s1.replace("--",'-')
    # df1['存货编码'] = df1['存货编码'].str.replace("（",'-').str.replace("）",'-').str.replace("(",'-').str.replace(")",'-')
    # df1['存货编码'] = df1['存货编码'].str.replace("--",'-')
    return s1

#处理最右边一个字符是'-'
def chuliRightHenggang(s):
    if s[-1] == '-':
        s1 = s[:-1]
    else :
        s1 = s
    return s1

# 处理最右边的字符为小写t，将其变更为大写的-T
def chuliRightT(s):
    if s[-2:] == '-t':
        s1 = s[:-2] + '-T'
    else :
        if s[-1] == 't':
            s1 = s[:-1] + '-T'
        else :
            s1 = s
    return s1

# 处理最右边的字符为小写-a，将其变更为大写的-A
def chuliRightA(s):
    if s[-2:] == '-a':
        s1 = s[:-2] + '-A'
    else :
        if s[-1] == 'a':
            s1 = s[:-1] + '-A'
        else :
            s1 = s
    return s1

#取消全部“-”，取最左边的
def delAllHenggang(s):
    s1 = s.split('-')[0]         
    s1 = s1.upper()              #全部改为大写
    return s1

def delrightA(s):
    if s.endswith('A'):
        s1  = s[:-1]
    else :
        s1 = s
    return s1

def chuliNumAfterHengGang(s):
    mat = re.search(pattern,s)
    if mat:
        str1 = mat.group(1)
        sl = int(mat.group(2))
    else :
        str1 = s
        sl = 1
    return str1,sl
def bianmaBiaozhun(s):
    # s = chuliFirstHenggang(s)
    s1 = chuliGangC(s)
    s1 = chuliXiaoxieA(s1)
    s1 = chuliXiaoxieT(s1)
    s1 = chuliTedingZhifu(s1)
    s1 = chuliHuohao(s1)
    s1 = chuliRightHenggang(s1)
    s1 = chuliRightT(s1)
    s1 = chuliRightA(s1)
    s1 = delAllHenggang(s1)
    # s1 = delrightA(s1)
    return s1



#对编码开头是78的前面直接加N0
def chuliQiba(s):
    if s.startswith('78'):
        s1 = 'N0' + s
    else :
        s1 = s
    return s1

def chuliXiaoshoubuDiaobojia(df_xiaoshou1):
    #将销售部价格中含字符的全部处理为数值
    df_xiaoshou1 = df_xiaoshou1.copy()
    df_xiaoshou1['内部调拨价'] = df_xiaoshou1['内部调拨价'].fillna(0)
    regex = re.compile(r'^(\d+)\D+')  
    df_xiaoshou1['内部调拨价'] = df_xiaoshou1['内部调拨价'].map(quShu)
    #将销售部存货标准处理成标准编码
    df_xiaoshou2 = df_xiaoshou1.copy()
    df_xiaoshou2['存货编码'] = df_xiaoshou2['存货编码'].map(chuliFirstHenggang)
    df_xiaoshou2['存货编码'] = df_xiaoshou2['存货编码'].map(bianmaBiaozhun)
    #将原编码和标准编码合并
    df_xiaoshou = pd.concat([df_xiaoshou1,df_xiaoshou2])
    df_xiaoshou.to_excel('销售部编码标准化.xlsx',index = False)
    #合并后生成的编码-调拨价字典
    bianma_diaobojia = dict(zip(df_xiaoshou.存货编码,df_xiaoshou.内部调拨价))
    return df_xiaoshou1,df_xiaoshou2,df_xiaoshou,bianma_diaobojia

def main():
    #销售部内部调拨价格汇总
    isno = easygui.boolbox('是否对销售部调拨价格进行汇总')
    df_xiaoshou1 = huizhongXiaoshoubuDiaobojia(isno)
    df_xiaoshou1,df_xiaoshou2,df_xiaoshou,bianma_diaobojia = chuliXiaoshoubuDiaobojia(df_xiaoshou1)
    #对聚水潭的主题分析文件进行处理
    #先处理横杠后带数字的   #组合编码，编码中含“+”
    pattern = r'(.*)-(\d{1,2})$'
    #聚水潭销售主题文件
    # fname_jusuitan = easygui.fileopenbox('请打开聚水潭销售主题文件')

    isno = easygui.boolbox('是否需要输入期间和平台')
    if isno:
        # fname_jusuitan = easygui.fileopenbox('请打开聚水潭各平台总销售')
        fname_jusuitan = r"F:\a00nutstore\008\zww08\002电商\聚水潭\聚水潭各平台销售\聚水潭各平台总销售.xlsx"
        df_jusuitan = pd.read_excel(fname_jusuitan,sheet_name = '原始数据',dtype = {'商品编码':str,'其他属性4':str})
        qijians = list(df_jusuitan.期间.unique())
        pingtais = list(df_jusuitan.平台.unique())
        pingtais = easygui.multchoicebox(msg = '请点先平台',choices = pingtais)
        qijians = easygui.multchoicebox(msg = '请点先期间',choices = qijians)
        conds = ''
        for pingtai in pingtais:
            condi1= f"(df_jusuitan['平台'] == '{pingtai}')"
            for qijian in qijians:
                condi2 =  f"(df_jusuitan['期间'] == '{qijian}')"
                condi = f'({condi1} & {condi2})'              #平台和期间的条件以“&”和连接
                if conds != '':
                    conds = conds +' | ' +condi               #各条件以"|"连接
                else :
                    conds = condi
        
        df_jusuitan = df_jusuitan[eval(conds)]
        df_jusuitan = df_jusuitan.assign(bianma1 = df_jusuitan.商品编码.apply(lambda x:chuliNumAfterHengGang(x)[0]))
        df_jusuitan = df_jusuitan.assign(shuliang = df_jusuitan.商品编码.apply(lambda x:chuliNumAfterHengGang(x)[1]))
          
    else :
        fname_jusuitan = easygui.fileopenbox('请打开聚水潭销售主题文件')
        df_jusuitan = pd.read_excel(fname_jusuitan,dtype = {'商品编码':str})
        df_jusuitan = df_jusuitan[~df_jusuitan.商品编码.isna()]
        df_jusuitan = df_jusuitan.assign(bianma1 = df_jusuitan.商品编码.apply(lambda x:chuliNumAfterHengGang(x)[0]))
        df_jusuitan = df_jusuitan.assign(shuliang = df_jusuitan.商品编码.apply(lambda x:chuliNumAfterHengGang(x)[1]))
    df_jusuitan.to_excel('jusuitan.xlsx',index = False)
    #爆炸，将编码中带
    df_jusuitan1 = df_jusuitan.copy()
    #对爆炸进行处理,增加一个过渡列"len",利用其长度计算爆炸后每个元素的数量
    df_jusuitan1['bianma2'] = df_jusuitan1['bianma1'].str.split('+')
    df_jusuitan1['len'] = df_jusuitan1['bianma2'].str.len()
    df_jusuitan1 = df_jusuitan1.assign(shuliang1 = np.where(df_jusuitan1.len == 1,df_jusuitan1['shuliang'] ,df_jusuitan1['shuliang']/df_jusuitan1['len']))
    df_jusuitan1 = df_jusuitan1.explode('bianma2')
    df_jusuitan1.to_excel('jusuitan1.xlsx',index = False)
    for i in ['实发金额','销售金额','销售毛利','净销售额','净销售成本','净销售毛利','基本金额','已付金额','优惠金额']:
        df_jusuitan1[f'{i}'] = np.where(df_jusuitan1.len == 1,df_jusuitan1[f'{i}'] ,df_jusuitan1[f'{i}']/df_jusuitan1['len'])
    
    df_jusuitan1.bianma2 = df_jusuitan1.bianma2.map(chuliQiba)
    df_jusuitan1.to_excel('jusuitan1.xlsx',index = False)
    df_jusuitan1.to_excel('聚水潭爆炸jusuitan1.xlsx',index = False)
    #爆炸后将聚水潭编码与大库编码匹配
    df_jusuitan2 = df_jusuitan1.copy()
    df_jusuitan2 = df_jusuitan2.assign(diaobojia = df_jusuitan2.bianma2.map(bianma_diaobojia))
    df_jusuitan2.to_excel('爆炸后编码标准化.xlsx',index = False)
    #对df_jusuitan2的编码bianma2
    jusuitan01 = df_jusuitan2[df_jusuitan2.diaobojia.isna()]
    jusuitan02 = df_jusuitan2[~df_jusuitan2.diaobojia.isna()]
    #将与大库编码匹配不上的编码进一步匹配
    #对聚水潭2的编码进行标准化
    jusuitan01['bianma3'] = jusuitan01['bianma2'].map(chuliFirstHenggang)
    jusuitan01['bianma3'] = jusuitan01['bianma3'].map(bianmaBiaozhun)
    jusuitan01.diaobojia = jusuitan01.bianma3.map(bianma_diaobojia)   #处理后再一次与大库匹配
    jusuitan02 = jusuitan02.assign(bianma3 = jusuitan02.bianma2)
    jusuitan01.to_excel('jusuitan01.xlsx',index = False)
    jusuitan02.to_excel('jusuitan02.xlsx',index = False)
    df_jusuitan3 = pd.concat([jusuitan01,jusuitan02])
    df_jusuitan3  = df_jusuitan3.assign(diaobojia1 = df_jusuitan3.diaobojia)
    df_jusuitan3.to_excel('编码处理后的聚水潭销售（未处理成本价）.xlsx',index = False)
    #对没有调拨价的，取聚水潭上的成本价
    jusuitan4 = df_jusuitan3.copy()
    jusuitan4.diaobojia1 = jusuitan4.diaobojia
    jusuitan4.diaobojia1 = jusuitan4.diaobojia1.fillna(0)
    jusuitan4 = jusuitan4.assign(diaobojia1 = np.where(jusuitan4.diaobojia1 == 0,jusuitan4.成本价,jusuitan4.diaobojia1))
    #经过上面处理，如果调拨价仍为0，则用净销售金额/实发数量
    #计算实际数量
    jusuitan4['净销量'] = jusuitan4['净销量'].fillna(0)
    jusuitan4['净销售成本'] = jusuitan4['净销售成本'].fillna(0)
    jusuitan4 = jusuitan4.assign(数量 = jusuitan4['净销量']*jusuitan4['shuliang1'])
    jusuitan4.diaobojia1 = jusuitan4.diaobojia1.fillna(0)
    def chuliChengben(d):
        if d['diaobojia1'] == 0 and d['数量'] != 0:
            d['diaobojia1'] = round(d['净销售额']/d['数量'],2)
        else :
            pass
        return d
    jusuitan4 = jusuitan4.apply(chuliChengben,axis = 1)
    jusuitan4.成本价 = jusuitan4.成本价.fillna(0)
    jusuitan4 = jusuitan4.assign(调拨成本 = jusuitan4['数量']*jusuitan4['diaobojia1'])
    jusuitan4 = jusuitan4.assign(净调拨毛利 = jusuitan4['净销售额'] - jusuitan4['调拨成本'])
    jusuitan4['净销售毛利'] = jusuitan4['净销售额'] - jusuitan4['净销售成本']
    jusuitan4 = jusuitan4.assign(收入 = round(jusuitan4['净销售额']/1.13,2))
    jusuitan4 = jusuitan4.assign(成本 = round(jusuitan4['调拨成本']/1.13,2))
    
    jusuitan4 = jusuitan4.rename(columns = {'bianma1': '编码1',
     'shuliang': '数量0',
     'bianma2': '编码2',
     'len': '元素',
     'shuliang1': '数量1',
     'diaobojia': '调拨价',
     'bianma3': '编码3',
     'diaobojia1': '调拨价1'})
    
    jusuitan4 = jusuitan4.assign(成本差异 = jusuitan4.成本价 - jusuitan4.调拨价1 )      #负数不正常
    jusuitan4.to_excel('编码处理后的聚水潭销售（已处理成本价）.xlsx',index = False)
    # os.startfile(fname)
    #聚水潭有而销售部没有调拨价
    jusuitan5  = jusuitan4.copy()
    jusuitan5.调拨价 = jusuitan5.调拨价.fillna(0)
    jusuitan5 = jusuitan5[jusuitan5.调拨价 ==  0]
    df_xiaoshouNo = pd.DataFrame(data = {'聚水潭':jusuitan5.商品编码,'编码':jusuitan5.编码3})
    df_xiaoshouNo.to_excel('聚水潭有而销售部没有调拨价.xlsx',index = False)
    #是否附加到原始数据中
    fname_jusuitan = r"F:\a00nutstore\008\zww08\002电商\聚水潭\聚水潭各平台销售\聚水潭各平台总销售.xlsx"
    isno = easygui.boolbox('是否将本期各平加工后销售数据添加到加工数据')
    if isno:
        df_0 = pd.read_excel(fname_jusuitan,sheet_name = '加工')
        max_row = df_0.shape[0]
        with pd.ExcelWriter(fname_jusuitan, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            jusuitan4.to_excel(writer, sheet_name='加工', startrow=max_row + 1, header=False, index=False)
    
    else :
        pass
    #经上面处理后，调拨成本比较准确，然后汇总各平台的销售
    columns = ['数量','净销售额', '净销售成本', '净销售毛利','运费收入','调拨成本','净调拨毛利', '收入', '成本']
    pivot = pd.pivot_table(jusuitan4,index = '平台',fill_value = None,aggfunc = 'sum',margins = True,margins_name = '合计', values = columns)
    pivot1 = pivot[columns]
    pivot1.to_excel('各平台销售收入毛利.xlsx') 


if __name__=='__main__':
    main()               
         