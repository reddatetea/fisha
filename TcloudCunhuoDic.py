#!/usr/bin/env python
# coding: utf-8

# In[16]:


'''
计算Tcloud中存货档货中存货含量字典20240917
'''
import pandas as pd
import re
import easygui
import openpyxl

def getCunhuoConcent(fname):
    pattern = r'(?P<num>\d+)本/件'
    regex = re.compile(pattern)
    df = pd.read_excel(fname,dtype = {'存货编码':'str'})
    def addBen(string):
        if string == '本':
            string =0
            return string
        else :
            mat = regex.search(string)
            
            if mat:
                string = int(mat.group('num'))
                return string
            else :
                string = 0
                return string
    df['计量单位1'] = df['计量单位'].map(addBen)
    content_dic = dict(zip(df['存货编码'],df['计量单位1'] ))
    return content_dic

def getCunhuoConcentFile(path,content_dic):
    riqi = easygui.enterbox('请输入日期')
    dicfile = f'存货档案字典{riqi}.xlsx'
    sheet_name = '存货档案'
    df_dic = pd.DataFrame.from_dict(content_dic,orient='index')
    df_dic = df_dic.reset_index()
    df_dic.columns = ['存货编码','含量']
    df_dic.to_excel(dicfile,sheet_name = sheet_name,index = False)
    return os.path.join(path,dicfile)
    
    
    

def main():
    fname  = easygui.fileopenbox('请点选存货档案')
    path,_ = os.path.split(fname)
    os.chdir(path)
    content_dic = getCunhuoConcent(fname)
    fname_dic = getCunhuoConcentFile(path,content_dic)
    os.startfile(fname_dic)
    
    


if __name__=='__main__':
    main()               
    
            
       
         
       
      
   

    


# In[ ]:




