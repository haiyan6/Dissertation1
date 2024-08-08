import pandas as pd

# read sheets
sheet1 = pd.read_excel('/workspaces/Dissertation1/Haiyan - Chinese Corporate Bond Data.xlsx', sheet_name='Corporate Bonds')
sheet2 = pd.read_excel('/workspaces/Dissertation1/Haiyan - Chinese Corporate Bond Data.xlsx', sheet_name='Daily Share Price & Market data')

# 假设时间列在第一个工作表的第一列 (列名为 'Date')
# 并且我们想选取第一个时间
target_time = pd.to_datetime(sheet1['Start Date'].iloc[0])

# 将第二个工作表的日期列转换为日期时间格式，假设列名为 'Date'
sheet2['Date'] = pd.to_datetime(sheet2['Name'])

# 找到第二个工作表中与目标时间匹配的行
target_row = sheet2[sheet2['Date'] == target_time]

# 如果找到了匹配的时间，获取目标时间前后10天的数据
if not target_row.empty:
    start_date = target_time - pd.Timedelta(days=10)
    end_date = target_time + pd.Timedelta(days=10)
    
    # 过滤出在此日期范围内的数据
    filtered_data = sheet2[(sheet2['Date'] >= start_date) & (sheet2['Date'] <= end_date)]
    
    print(filtered_data)
else:
    print("在第二个工作表中未找到匹配的时间。")
