import openpyxl
import re
keynames = ['晴川自制肉圆/500g', '晴川自制鱼圆/500g', '老武汉藕夹/500g', '荆州鱼糕/500g', '晴川秘制酱牛腱子/500g', '招牌葱油鸡/一只', '梅菜扣肉/份', '招牌羊肉火锅/500g', '晴川珍珠肉圆/500g', '糍粑鱼/500g', '晴川自制炸鱼块/500g', '晴川肉糕/500g', '羊城烧鸭/半只']
print(keynames[0])
# wb1 = openpyxl.Workbook()
# ws1 = wb1.active
fuhao = ';\/?*[]'
for j in fuhao:
    if j in keynames[0]:
        sheetname = keynames[0].replace(j,'')
    else :
        continue
print(sheetname)

# a = '晴川自制肉圆/500g'
# b= a.replace('/','')
# print(b)

    #sheetname = keynames[0].replace(j,'')
# print(keynames[0])
# ws1.title = sheetname
# wb1.save(sheetname)


result00
#%%
(
    pd.concat([df_2022, df_2023], keys=['2022', '2023'], names=['year', 'idx'])
    .reset_index()
    .assign(years=lambda x: x.groupby('a').year.transform(lambda y: [y.to_list()]*len(y)))
    .style
    .apply(lambda x: ['color: red' if x.years == ['2022'] else '']*len(x), axis=1)
    .apply(lambda x: ['background-color: #73BF00' if x.years == ['2023'] else '']*len(x), axis=1)
    .hide('years', axis=1)
)