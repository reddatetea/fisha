(
      pd.concat([df_2022, df_2023], keys=['2022',  '2023'], names=['year',  'idx'])
      .reset_index()
      .assign(years=lambda  x: x.groupby('a').year.transform(lambda  y: [y.to_list()]*len(y)))
      .style
      .apply(lambda  x: ['color: red'  if  x.years == ['2022']  else  '']*len(x), axis=1)
      .apply(lambda  x: ['background-color: #73BF00'  if  x.years == ['2023']  else  '']*len(x), axis=1)
      .hide('years', axis=1)
)