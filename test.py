'''
hello git
'''
# 使用链式方法一次性完成所有操作
(
      df
       # 创建季度周期索引
      .assign(quarter=lambda  x: pd.PeriodIndex(x['date'], freq='Q'))
       # 按季度分组并计算总销售额
      .groupby('quarter')
      .agg(quarterly_sales=('amount',  'sum'))
       # 计算环比增长率
      .assign(qoq_growth=lambda  x: x['quarterly_sales'].pct_change() *  100)
       # 格式化增长率显示为百分比
      .assign(qoq_growth=lambda  x: x['qoq_growth'].round(2))
       # 重置索引使季度成为一列
      .reset_index()
       # 重命名列
      .rename(columns={'quarter':  '季度',  
                            'quarterly_sales':  '季度销售额',
                            'qoq_growth':  '环比增长(%)'})
)