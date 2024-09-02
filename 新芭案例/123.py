from io import StringIO
import pandas as pd
import numpy as np
import re
from sqlalchemy import true

data = '''
�鵥,����,ISBN
�ʱ��� 1,Ԭĳ,23445-2342
�ʱ��� 2,Ԭĳ,23445-2342
�ʱ��� 3,Ԭĳ,23445-2342
��ѧ���� ��,����,a25245-32425
��ѧ���� ��,����,a25245-32425
���ݽṹ,��ĳ,x342w-ssa
'''
def mark(ser:pd.Series):
    if ser.isna().all():
        return  ''
    elif ser.str.isdecimal().all():
        return f'{ser.min()}-{ser.max()}'
    else :
        return ser.sum()

df = pd.read_csv(StringIO(data))
df1 = df.�鵥.str.split(' ',expand=True)
df1 = df1.set_axis(['����','��ע'],axis=1)
new_df = pd.concat([df,df1],axis=1)
result = new_df.groupby('����',as_index=False,sort=False).agg({'����':max,'ISBN':max,'��ע':mark})
result
