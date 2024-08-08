# pip install openpyxl

import pandas as pd

# read sheets
sheet1 = pd.read_excel('/workspaces/Dissertation1/Haiyan - Chinese Corporate Bond Data.xlsx', sheet_name='Corporate Bonds')
sheet2 = pd.read_excel('/workspaces/Dissertation1/Haiyan - Chinese Corporate Bond Data.xlsx', sheet_name='Daily Share Price & Market data')
ewl=10
ewr=10

# 将 'Date' 列转换为 datetime 格式
sheet1['Date'] = pd.to_datetime(sheet1['Start Date'])
# 过滤掉2024年以后的数据
sheet1 = sheet1[sheet1['Date'] <= '2024-12-31']

# 我们想选取第一个时间
target_time = pd.to_datetime(sheet1['Start Date'].iloc[0])

# 删除第二个工作表的第一行
sheet2 = sheet2.iloc[1:].reset_index(drop=True)

# 将第二个工作表的日期列转换为日期时间格式，假设列名为 'Date'
sheet2['Date'] = pd.to_datetime(sheet2['Name'])

# 找到第二个工作表中与目标时间匹配的行
target_index = sheet2[sheet2.iloc[:, 0] == target_time].index

# 如果找到了匹配的时间
if not target_index.empty:
    # 取第一个匹配的索引
    target_index = target_index[0]
    
    # 确定前后10个单元格的范围
    start_index = max(target_index - ewl, 0)
    end_index = min(target_index + ewr + 1, len(sheet2))  # +1 因为范围是闭区间
    
    # 选取数据
    selected_data = sheet2.iloc[start_index:end_index]
    
    print(selected_data)
else:
    print("在第二个工作表中未找到匹配的时间。")