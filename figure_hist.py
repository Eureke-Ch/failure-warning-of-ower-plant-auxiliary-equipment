import xlrd
import matplotlib.pyplot as plt
import math
import numpy as np
from pylab import *
dataCeshi = []
data_xsls = xlrd.open_workbook("./data/14A17Process.xlsx") #打开此地址下的exl文档
sheet_name = data_xsls.sheets()[0]
count_nrows = sheet_name.nrows  
a = []
for i in range(1,count_nrows,40):
    if i < 3000 or (i>250000 and i<300000) or(i>380000 and i<490000) or (i > 760000 and i< 820000):# or (i>320000 and i<360000) or (i>840000 and i<990000):
        pass
    else:
        if sheet_name.cell_value(i,8)>7 or sheet_name.cell_value(i,8)<-7:
            pass
        else:
            a.append((sheet_name.cell_value(i,6)))
dataCeshi.append(a)

def normfun(x,mu,sigma):
    pdf = np.exp(-((x - mu)**2)/(2*sigma**2)) / (sigma * np.sqrt(2*np.pi))
    return pdf
#data = dataCeshi[0].sort()
data = np.array(dataCeshi[0])
data.sort()
mu =np.mean(data) #计算均值 
sigma =np.std(data) 
y = normfun(data, mu, sigma)
plt.plot(data,y,color = 'red',linewidth = '2')

font1 = {
'weight' : 'normal',
'size'   : 20,
}

mpl.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus']=False
plt.xlabel('一次风机X方向上的轴承振动',font1)
plt.ylabel('频率',font1)
plt.hist(dataCeshi[0],bins = 200,density = True,color = 'cyan')
plt.show()

'''
fig = plt.figure()
ax = []
for i in range(9):
    ax.append(fig.add_subplot(3,3,i+1))
for i in range(9):
    ax[i].hist(dataCeshi[i],bins=400,density = True,color = 'red')

plt.show()

'''


