import xlrd
import matplotlib.pyplot as plt
from pylab import *
import matplotlib.dates as mdates
import mpl_toolkits.axisartist as axisartist
def read_xsls(xlsx_path):
    data_xsls = xlrd.open_workbook(xlsx_path)
    sheet_name = data_xsls.sheets()[0] 
    count_nrows = sheet_name.nrows 
    for i in range(3000,count_nrows,100):
        data1.append(sheet_name.cell_value(i,2))
        data2.append(sheet_name.cell_value(i,8))
        X.append(xlrd.xldate_as_tuple(sheet_name.cell_value(i,0),0))

data1 = [];data2 = [];X = []
data_path1 = './data/14A17.xlsx'
#data_path2 = './data/14A17-7-12.xlsx'
read_xsls(data_path1)
#read_xsls(data_path2)
mpl.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus']=False

for i in range(len(data2)):
    X[i] = datetime.datetime(*X[i])

font1 = {
'weight' : 'normal',
'size'   : 14,
}
plt.xlabel('日 期',font1)
plt.ylabel('一次风机电流/(A)',font1)
plt.ylim((-200, 350))
plt.xticks(rotation = 20)
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
plt.plot(X,data1,'k',linewidth = '0.4')
#plt.title('14号机组一次风机A2017年处理后运行数据',font1)
plt.show()

'''
plt.subplot(2,1,1)
plt.xlabel('时间')
plt.ylabel('一次风机电动机前轴承温度/℃')
plt.xticks(rotation = 20)
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
plt.plot(X,data1,'k',linewidth = '0.3')
plt.subplot(2,1,2)
plt.xlabel('时间')
plt.ylabel('一次风机X方向上的轴承振动/(mm/s)')
plt.xticks(rotation = 20)
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
plt.plot(X,data2,'k',linewidth = '0.3')
plt.suptitle('14A一次风机2017年运行数据')
plt.show()
'''
