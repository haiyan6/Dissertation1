# pip install openpyxl

import pandas as pd
import numpy as np
import scipy as sp
from scipy . optimize import leastsq
import matplotlib.pyplot as plt

# read sheets
CBG = pd.read_excel('/workspaces/Dissertation1/Haiyan - Chinese Corporate Bond Data.xlsx', sheet_name=[0,1])
# sheet2 = pd.read_excel('/workspaces/Dissertation1/Haiyan - Chinese Corporate Bond Data.xlsx', sheet_name='Market & Stock Premium')
ewl=10
ewr=10
def error(params, x, y):
    a, b = params
    return y - (a * x + b)

i=0
row_num_x = list(CBG[0]['Code']).index(CBG[1]['Start Date'][i])
x = CBG[0]['CHSASHR'][row_num_x - 300: row_num_x - 49]
y = CBG[0][CBG[1]['Code'][i]][row_num_x - 300: row_num_x - 49]
p0 = [0, 1]
Para = leastsq(error, p0, args=(x, y))
a, b = Para[0]
print(a)
print(b)
y_pred = a * x + b
residuals = y - y_pred
print("cancha", residuals)
plt.figure(figsize=(10, 6))
plt.scatter(x, residuals, color='blue', label='Residuals')
plt.hlines(0, xmin=x.min(), xmax=x.max(), colors='red', linestyles='dashed')
plt.xlabel('x')
plt.ylabel('Residuals')
plt.title('Residuals Plot')
plt.legend()
plt.show()
# for i in range(len(CBG[1]['Code'])):
#     row_num_x = list(CBG[0]['Code']).index(CBG[1]['Start Date'][i])
#     x = CBG[0]['CHSASHR'][row_num_x - 300: row_num_x - 49]
#     y = CBG[0][CBG[1]['Code'][i]][row_num_x - 300: row_num_x - 49]
#     p0 = [0, 1]
#     Para = leastsq(error, p0, args=(x, y))
#     a, b = Para[0]




# # 将 'Date' 列转换为 datetime 格式
# sheet1['Date'] = pd.to_datetime(sheet1['Start Date'])


# # 我们想选取第一个时间
# target_time = pd.to_datetime(sheet1['Start Date'].iloc[0])

# # 删除第二个工作表的第一行
# sheet2 = sheet2.iloc[1:].reset_index(drop=True)

# # 将第二个工作表的日期列转换为日期时间格式，假设列名为 'Date'
# sheet2['Date'] = pd.to_datetime(sheet2['Name'])

# # 找到第二个工作表中与目标时间匹配的行
# target_index = sheet2[sheet2.iloc[:, 0] == target_time].index

# # 如果找到了匹配的时间
# if not target_index.empty:
#     # 取第一个匹配的索引
#     target_index = target_index[0]
    
#     # 确定前后10个单元格的范围
#     start_index = max(target_index - ewl, 0)
#     end_index = min(target_index + ewr + 1, len(sheet2))  # +1 因为范围是闭区间
    
#     # 选取数据
#     selected_data = sheet2.iloc[start_index:end_index]
    
#     print(selected_data)
# else:
#     print("在第二个工作表中未找到匹配的时间。")