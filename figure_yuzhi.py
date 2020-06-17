from numpy import *
import math
import xlrd
import numpy as np
import openpyxl
import datetime
import matplotlib.pyplot as plt
from pylab import *
import matplotlib.dates as mdates
import mpl_toolkits.axisartist as axisartist
def norm_pdf_multivariate(x, mu, sigma):
    size = len(x)
    if size == len(mu) and (size, size) == sigma.shape:
        det = linalg.det(sigma)
        if det == 0:
            raise NameError("The covariance matrix can't be singular")

        norm_const = 1.0/ (math.pow((2*pi),float(size)/2) * math.pow(det,1.0/2) )
        x_mu = matrix(x - mu)
        inv = sigma.I        
        result = math.pow(math.e, -0.5 * (x_mu * inv * x_mu.T))
        return norm_const * result
    else:
        raise NameError("The dimensions of the input don't match")

#读取清洗后的数据

datasigma = [];datatimeYes = []
data_xsls = xlrd.open_workbook("./data/14A17Process.xlsx") #打开此地址下的exl文档
#data_xsls = xlrd.open_workbook("13ANew.xls") 
sheet_name = data_xsls.sheets()[0]  
count_nrows = sheet_name.nrows  
for i in range(1,count_nrows,10):
    if i < 3000 or (i>250000 and i<300000) or(i>380000 and i<490000) or (i > 760000 and i< 820000):
    #if i < 3000 or (i>260000 and i<290000) or (i > 390000 and i< 480000) or (i>770000 and i<810000):
    #if i < 3000 or (i>50000 and i<100000) or (i > 330000 and i< 350000) or (i>850000 and i<990000):
        pass
    else:
        b = []
        for j in range(9):
            b.append(sheet_name.cell_value(i,j+1))
        datasigma.append(b)
        a = xlrd.xldate_as_tuple(sheet_name.cell_value(i,0),0)
        datatimeYes.append(datetime.datetime(*a))

#生成均值和协方差矩阵
sigma = matrix(cov(array(datasigma).T))
mean = []
for i in range(9):
    l = len(datasigma)
    mean.append(sum(datasigma[j][i]for j in range(l))/l)
mu = array(mean)

data = [];datatime = []
data_xsls = xlrd.open_workbook("./data/14A17Process.xlsx") #打开此地址下的exl文档
#data_xsls = xlrd.open_workbook("13ANew.xls") 
sheet_name = data_xsls.sheets()[0]  
count_nrows = sheet_name.nrows  
for i in range(380000,405000):
    b = []
    for j in range(9):
        b.append(sheet_name.cell_value(i,j+1))
    data.append(b)
    a = xlrd.xldate_as_tuple(sheet_name.cell_value(i,0),0)
    datatime.append(datetime.datetime(*a))

Y = []
for i in range(len(data)):
    a = norm_pdf_multivariate(data[i], mu, sigma)
    if a > 1*pow(10,-7):
        a = 1*pow(10,-7)
    Y.append(a)

mpl.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus']=False
font1 = {
'weight' : 'normal',
'size'   : 12,
}

plt.xlabel('日 期',font1)
plt.ylabel('多元高斯分布计算得到的P(X)',font1)
plt.xticks(rotation = 20)
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
plt.scatter(datatime,Y,marker='o',s=3,color = 'b',label='P(Xi)')
Z = [5*pow(10,-8) for i in range(len(Y))]
plt.plot(datatime,Z,marker='o',markersize=2,color = 'r',label='阈值ε')
plt.legend()
#plt.title('14号机组一次风机A2017年处理后运行数据',font1)
plt.show()



