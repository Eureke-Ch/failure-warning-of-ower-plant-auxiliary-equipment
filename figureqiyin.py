import xlrd
import matplotlib.pyplot as plt
from pylab import *
def read_xsls(xlsx_path):
    data_xsls = xlrd.open_workbook(xlsx_path)
    sheet_name = data_xsls.sheets()[0] 
    count_nrows = sheet_name.nrows 
    for i in range(3,count_nrows):
        for j in range(6):
            data[j].append(sheet_name.cell_value(i,j+1))



data = [[]for i in range(6)]
data_path1 = './data/14炉汽引风机18-1-12.xlsx'
#data_path2 = './data/14炉汽引风机17-7-12.xlsx'
#data_path3 = './数据/14A18-1-6.xlsx'
#data_path4 = './数据/14A18-7-11.xlsx'
read_xsls(data_path1)
#read_xsls(data_path2)
#read_xsls(data_path3)
#read_xsls(data_path4)
X = []
for i in range(len(data[0])):
    X.append(i*12/len(data[0])+1)
mpl.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus']=False
fig = plt.figure()
ax = []
for i in range(6):
    ax.append(fig.add_subplot(2,3,i+1))
ax[0].set_ylabel('DOXAA31CS18')
ax[0].plot(X,data[0])
ax[1].set_ylabel('前轴承温度3')
ax[1].plot(X,data[1])
ax[2].set_ylabel('中轴承温度3')
ax[2].plot(X,data[2])
ax[3].set_ylabel('后轴承温度3')
ax[3].plot(X,data[3])
ax[4].set_ylabel('轴承振动X方向')
ax[4].plot(X,data[4])
ax[5].set_ylabel('轴承振动Y方向')
ax[5].plot(X,data[5])

plt.suptitle('14B汽引风机2018-1-12月份运行状态')
plt.show()