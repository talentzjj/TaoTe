import pandas as pd
import numpy as np


#===================================== inpu args =====================================

file_name1 = '业务管理-淘特(伙聚版)-BD信息下载-2022-07-04.xlsx'
file_name2 = '0703-伙聚数据.xlsx'
file_name3 = '行为数据_220704_103823_120.xlsx'
file_name4 = '质量等级2213994253594_220704_31011.xlsx'

yesterday = pd.Timestamp('2022-07-03')

output_name1 = '0703用户行为.xlsx'
output_name2 = '0703行为数据.xlsx'

#===================================== read =========================================

data_dir = 'data/'

data1 = pd.read_excel(pd.ExcelFile(data_dir+file_name1))
data2 = pd.read_excel(pd.ExcelFile(data_dir+file_name2))
data3 = pd.read_excel(pd.ExcelFile(data_dir+file_name3))
data4 = pd.read_excel(pd.ExcelFile(data_dir+file_name4))

data1.rename(columns={'官方ID':'业务员id'}, inplace=True)
data2.rename(columns={'拉新业务员ID':'业务员id'}, inplace=True)
data4.rename(columns={'业务员ID':'业务员id'}, inplace=True)


#===================================== table-1 =====================================

d1 = data2[(data2['拉新状态']!='拉新失败') & (data2['是否零钱包支付']=='否') & (data2['30日内是否退款']=='否')]
d1['日期'] = pd.to_datetime(d1['日期'])
d1 = d1[d1['日期'] == yesterday]

D1 = pd.DataFrame(columns=['日期','业务员id', '是否参与活动','当日是否引导领首月红包','当日是否参与养小鸡','当日是否参与签到领红包',
           '当日是否参与天天一元购','当日是否参与现金签到','当日是否成功添加到桌面', '作业数量','一级渠道','业务员姓名'])

IDs = d1['业务员id'].unique()
for i in range(len(IDs)):
    D1.loc[i, '业务员id'] = IDs[i]
    Di = d1[d1['业务员id'] == IDs[i]]

    D1.loc[i, '一级渠道'] = Di['一级渠道'].iloc[0]
    D1.loc[i, '日期'] = Di['日期'].iloc[0]
    D1.loc[i, '业务员姓名'] = Di['业务员姓名'].iloc[0]

    D1.loc[i, '当日是否引导领首月红包'] = (Di['当日是否引导领首月红包'] == '是').sum()
    D1.loc[i, '当日是否参与养小鸡'] = (Di['当日是否参与养小鸡'] == '是').sum()
    D1.loc[i, '当日是否参与签到领红包'] = (Di['当日是否参与签到领红包'] == '是').sum()
    D1.loc[i, '当日是否参与天天一元购'] = (Di['当日是否参与天天一元购'] == '是').sum()
    D1.loc[i, '当日是否参与现金签到'] = (Di['当日是否参与现金签到'] == '是').sum()
    D1.loc[i, '当日是否成功添加到桌面'] = (Di['当日是否成功添加到桌面'] == '是').sum()
    D1.loc[i, '是否参与活动'] = (Di['是否参与活动'] == '是').sum()
    D1.loc[i, '作业数量'] = Di.shape[0]

D1['日期'] = D1['日期'].dt.strftime('%Y-%m-%d')
D1['业务员id'] = D1['业务员id'].astype(str)

D1.rename(columns={'业务员id':'业务员ID'}, inplace=True)

D1.to_excel(data_dir+output_name1, sheet_name='sheet1', index=False)

#===================================== table-2 =====================================

d2 = data3[data3['跨地区作业比例'].notnull()]
d2['日期'] = pd.to_datetime(d2['日期'], format='%Y%m%d')
d2 = d2[d2['日期'] == yesterday]

d2['一级渠道'] = np.NaN
d2['所属渠道'] = np.NaN
d2['零钱包付款比例'] = np.NaN
d2['首登比例'] = np.NaN
d2['退款率']= np.NaN
d2['规范率']= np.NaN

d3 = data2[(data2['拉新状态']!='拉新失败')]
d3['日期'] = pd.to_datetime(d3['日期'])
d3 = d3[d3['日期'] == yesterday]

for i,r in d2.iterrows():
    ID = d2.loc[i, '业务员id']
    d = data1[data1['业务员id'] == ID]
    d2.loc[i, '所属渠道'] = d['所属渠道'].iloc[0]
    d2.loc[i, '一级渠道'] = d['一级渠道'].iloc[0]

    d = d3[d3['业务员id']==ID]
    round(d2.loc[i, '零钱包付款比例'] = (d['是否零钱包支付'] == '是').sum() / d['是否零钱包支付']，4）.shape[0]
    d2.loc[i, '首登比例'] = (d['是否新登'] == '是').sum() / d['是否新登'].shape[0]
    d2.loc[i, '退款率'] = (d['30日内是否退款'] == '是').sum() / d['30日内是否退款'].shape[0]
    d2.loc[i, '规范率'] = (d['是否规范操作'] == '是').sum() / d['是否规范操作'].shape[0]

D2 = d2
D2['零钱包付款比例'] = D2['零钱包付款比例'].map(lambda x: format(x,'.2%'))
D2['首登比例'] = D2['首登比例'].map(lambda x: format(x,'.2%'))
D2['退款率'] = D2['退款率'].map(lambda x: format(x,'.2%'))
D2['规范率'] = D2['规范率'].map(lambda x: format(x,'.2%'))
D2['日期'] = D2['日期'].dt.strftime('%Y-%m-%d')

D2['业务员id'] = D2['业务员id'].astype(str)
D2.rename(columns={'业务员id':'业务员ID'}, inplace=True)
D2.rename(columns={'录入渠道':'所属渠道'}, inplace=True)
D2.to_excel(data_dir+output_name2, sheet_name='sheet1', index=False)

print("Finished")

