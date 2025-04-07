import re
import pandas as pd
import os
import openpyxl
import easygui

path0 = easygui.diropenbox('请点选包含"hello"文件夹')
# path0 = r"l:\fast"
lst = os.listdir(path0)
lst1 = [i  for i in lst if i.startswith('hellorbsq') ]   #  and i.startswith('hellorbsq')
data = []
# data = [['path','file']]
for i in lst1:
    path1 = os.path.join(path0,i)
    lst2 = os.listdir(path1)
    for j in lst2:
        path2 = os.path.join(path1,j)
        data.append([i,j])
df = pd.DataFrame(data,columns = ['path','file'])
fname = r'F:\a00nutstore\008\zww08\hellorbsq\hellorbsqmingxi.xlsx'
wb = openpyxl.load_workbook(fname)
ws = wb.active
max_row = ws.max_row
with pd.ExcelWriter(fname, engine='openpyxl',mode='a', if_sheet_exists='overlay')  as writer:
    df.to_excel(writer, sheet_name = ws.title,startrow = max_row,index = False,header = None)
os.startfile(fname)



# In[ ]:





# In[ ]:




