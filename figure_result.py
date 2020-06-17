import datetime
import matplotlib.pyplot as plt
from pylab import *
import matplotlib.dates as mdates
import mpl_toolkits.axisartist as axisartist
import xlrd
data = [];datatime = [];Y1 = [];Y2 = []
data_xsls = xlrd.open_workbook("clusting.xlsx") #打开此地址下的exl文档
sheet_name = data_xsls.sheets()[0]  
count_nrows = sheet_name.nrows  
for i in range(1,count_nrows):
    datatime.append(i-1)
    if sheet_name.cell_value(i,10) == 0:
        Y1.append(1)
        Y2.append(None)
    else:
        Y2.append(0)
        Y1.append(None)



mpl.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus']=False
font1 = {
'weight' : 'normal',
'size'   : 12,
}
plt.xlabel('样本序号',font1)
plt.ylabel('模型预警结果',font1)

plt.plot(datatime,Y1,marker='o',markersize=4,color = 'r',label='故障')
plt.plot(datatime,Y2,marker='o',markersize=4,color = 'b',label='正常')
plt.legend()
#plt.title('14号机组一次风机A2017年处理后运行数据',font1)
plt.show()