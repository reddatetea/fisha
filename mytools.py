'''
mytools
'''
import os
import re
import openpyxl
import numpy as np
import pandas as pd
import xlwings as xw
import easygui
from tkinter import *
from collections import ChainMap
import functools
import time

def getTotal(df):                       #根据df每列数据类型，计算合计数，只计算float,float64,int三种类型
    dic1 = dict(df.dtypes)
    total = []
    for i in df.columns.to_list():
        dtype = dic1.get(i)
        if dtype in ['float64','int','float']:
            total0 = df[i].sum()
            total.append(total0)
        else:
            total0 = ''
            total.append(total0)
    return total

def addHuizhong(df):      #df加舍计数
    total = getTotal(df)
    total[0] = '合计'
    dic = dict(zip(df.columns,total))
    df1 = pd.DataFrame([dic])
    result = pd.concat([df,df1])
    # df.loc['合计'] = total
    return result

def getStartriqiEndriqi():                  #获取开始日期和结束日期（字符串格式和日期格式）
    win = Tk()
    win.title('请输入日期，格式如"2023-1-12"')
    label1 = Label(win,text='开始日期',font=14)
    label1.grid(pady=10,row=0,column=0)
    text_var = StringVar()
    text1 = Entry(win, width=50, textvariable=text_var)
    text1.grid(row=0,column=1)
    
    label2 = Label(win,text='结束日期',font=14)
    label2.grid(pady=10,row=2,column=0)
    text_var = StringVar()
    text2 = Entry(win, width=50, textvariable=text_var)
    text2.grid(row=2,column=1)
    def submit():
        global start_riqi_str
        global end_riqi_str
        global start_riqi
        global end_riqi
        start_riqi_str = text1.get()
        end_riqi_str = text2.get()
        start_riqi = pd.Timestamp(start_riqi_str)
        end_riqi = pd.Timestamp(end_riqi_str)
    
       
    Button(win,text="确定起止日期",command=submit).grid(row=5,column=0)
   
    # start_riqi = pd.Timestamp(start_riqi)
    
    win.mainloop()
    return start_riqi_str,end_riqi_str,start_riqi,end_riqi

def gameIsOver():
    win=Tk()
    win.title("恭喜")
    win.geometry("300x100")
    text=Label(win,text="\n程序结序!",font="Times 18 bold").pack()
    win.mainloop()

def choiceOneFromTwo():
    def result1():
        if v.get() == 1:
            re.config(text = '您选择的是"汇总开具"')
            
        else:
            re.config(text = '您选择的是"单张开具"')
        return v.get()
            
    # from tkinter import *
    win = Tk()
    win.title("选择发票开具方式")   #设置窗口标题
    win.geometry("300x150")   #设置窗口大小
    text = Label(win, text="汇总开具 或 单张开具",font="14").pack(anchor=W)
    # 该变量绑定单选按钮的值
    v = IntVar()
    ans1=Radiobutton(win, text="汇总开具", variable=v, value=1,font="12",selectcolor="#F1D4C9")
    ans1.pack(anchor=W)
    ans2=Radiobutton(win, text="单张开具", variable=v, value=2,font="12",selectcolor="#F1D4C9")
    ans2.pack(anchor=W)
    button = Button(win, text="提交", command=result1,font="14",bg="#F1C57E",relief="groove").pack()
    re = Label(win)     #显示答案的文本框
    re.pack()
    win.mainloop()
    return v.get()

def huizhongJener(df):
    pivot = pd.pivot_table(df,index = '发票抬头',values = ['商品数量','发票金额'],aggfunc = 'sum')
    shuliang = pivot['商品数量']
    jinger = pivot['发票金额']
    df1['商品数量'] = df1['发票抬头'].map(shuliang)
    df1['发票金额'] = df1['发票抬头'].map(jinger)
    df2 = df1.drop_duplicates(subset = '发票抬头',keep = 'last')  #去重
    return df2

def huizhongJener1(df_pdd):
    #非个人
    df_pdd_feigeren = df_pdd.loc[df_pdd['发票抬头'] != '个人']
    #个人
    df_pdd_geren = df_pdd.loc[~(df_pdd['发票抬头'] != '个人')]
    pivot = pd.pivot_table(df_pdd_feigeren,index = '发票抬头',values = ['商品数量','发票金额'],aggfunc = 'sum')
    shuliang = pivot['商品数量']
    jinger = pivot['发票金额']
    #获取发票抬头_订单号 字典
    gp = df_pdd_feigeren.groupby('发票抬头',sort = False,as_index = False)
    data = []
    for k,v in gp:
        v1 =[k,','.join(v['订单号'].to_list())]
        data.append(v1)
    df_  = pd.DataFrame(data,columns = ['发票抬头','订单号'])
    dingdans = dict(zip(df_['发票抬头'],df_['订单号']))
    df_pdd_feigeren['商品数量'] = df_pdd_feigeren['发票抬头'].map(shuliang)
    df_pdd_feigeren['发票金额'] = df_pdd_feigeren['发票抬头'].map(jinger)
    df_pdd_feigeren['订单号'] = df_pdd_feigeren['发票抬头'].map(dingdans)
    
    df_pdd_feigeren = df_pdd_feigeren.drop_duplicates(subset = '发票抬头',keep = 'last')  #去重
    df_pdd = pd.concat([df_pdd_feigeren,df_pdd_geren])
    
    return df_pdd

def inputText(input = '请输入传入的字符'):      #输入字符
    def content():
        global input_str
        input_str = str.get()
        print(input_str)
        
    
    root = tk.Tk()
    root.title = 'input_str'
    str = tk.StringVar()
    #设置默认文本
    str.set(input)
    tk.Entry(root, textvariable=str,justify="center").place(x=20, y=8)
    tk.Button(root, text="确认并关闭本窗口", command=content).place(x=40, y=40)
    root.mainloop()
    return input_str

def dicPiliang(lst,data_type):
    dic = {}
    for i in lst:
        exec("dic[f'{i}'] = data_type")  
    return dic


def format_conversion(styler):
     return (styler.set_properties(**{'text-align': 'right'})
                   .format({'conversion': '{:.1%}'}))

def xlsToXlsx(file_name):
    '''
    将xls格式转换为xlsx格式
    '''
    # file_name = '2222.xls'
    new_name = file_name+'x'

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    wb = app.books.open(file_name)   # 打开现有excel
    wb.api.SaveAs(new_name, 51)      # 参数 51 为xlsx格式。56为 Excel 97-2003的xls版本
    app.kill()     # 使用kill()关闭进程
    return new_name


def clock(func):
    @functools.wraps(func)
    def clocked(*args, **kwargs):
        t0 = time.perf_counter()
        result = func(*args, **kwargs)
        elapsed = time.perf_counter() - t0
        name = func.__name__
        arg_lst = [repr(arg) for arg in args]
        arg_lst.extend(f'{k}={v!r}' for k, v in kwargs.items())
        arg_str = ', '.join(arg_lst)
        print(f'[{elapsed:0.8f}s] {name}({arg_str}) -> {result!r}')
        return result
    return clocked






def main():
    # 调用
    df = pd.DataFrame({'trial': list(range(5)),
                    'conversion': [0.75, 0.85, np.nan, 0.7, 0.72]})
    (df.style
        .highlight_min(subset=['conversion'], color='green')
        .pipe(format_conversion)
        .set_caption("Results with minimum conversion highlighted.")
    )


if __name__=='__main__':
    main()     
    





