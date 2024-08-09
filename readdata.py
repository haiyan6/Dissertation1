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

# i=0
# row_num_x = list(CBG[0]['Code']).index(CBG[1]['Start Date'][i])
# x = CBG[0]['CHSASHR'][row_num_x - 300: row_num_x - 49]
# y = CBG[0][CBG[1]['Code'][i]][row_num_x - 300: row_num_x - 49]
# p0 = [0, 1]
# Para = leastsq(error, p0, args=(x, y))
# a, b = Para[0]
# print(a)
# print(b)
# y_pred = a * x + b
# residuals = y - y_pred
# print("cancha", residuals)
# plt.figure(figsize=(10, 6))
# plt.scatter(x, residuals, color='blue', label='Residuals')
# plt.hlines(0, xmin=x.min(), xmax=x.max(), colors='red', linestyles='dashed')
# plt.xlabel('x')
# plt.ylabel('Residuals')
# plt.title('Residuals Plot')
# plt.legend()
# plt.show()



for i in range(len(CBG[1]['Code'])):
    row_num_x = list(CBG[0]['Code']).index(CBG[1]['Start Date'][i])
    x = CBG[0]['CHSASHR'][row_num_x - 300: row_num_x - 49]
    y = CBG[0][CBG[1]['Code'][i]][row_num_x - 300: row_num_x - 49]
    p0 = [0, 1]
    Para = leastsq(error, p0, args=(x, y))
    a, b = Para[0]



