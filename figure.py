import xlrd
import matplotlib.pyplot as plt
from pylab import *
import matplotlib.dates as mdates
import mpl_toolkits.axisartist as axisartist
def read_xsls(xlsx_path):
    data_xsls = xlrd.open_workbook(xlsx_path)
    sheet_name = data_xsls.sheets()[0] 
    count_nrows = sheet_name.nrows 
    for i in range(1,count_nrows,20):
        X.append(xlrd.xldate_as_tuple(sheet_name.cell_value(i,0),0))
        for j in range(9):
            data[j].append(sheet_name.cell_value(i,j+1))


X = []
data = [[]for i in range(9)]
data_path1 = './data/14A17clean.xlsx'
data_path2 = './data/14A18clean.xlsx'
# data_path3 = './data/14A18-1-6.xlsx'
# data_path4 = './data/14A18-7-11.xlsx'
read_xsls(data_path1)
read_xsls(data_path2)
# read_xsls(data_path3)
# read_xsls(data_path4)
for i in range(len(data[0])):
    X[i] = datetime.datetime(*X[i])

mpl.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus']=False

fig = plt.figure()
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
plt.subplot(331)
plt.ylim((-30, 100))
plt.xticks(rotation = 20,fontsize=6)
plt.ylabel('调节挡板阀位')
plt.plot(X,data[0],linewidth = '0.6')
plt.subplot(332)
plt.ylim((-200, 600))
plt.xticks(rotation = 20,fontsize=6)
plt.ylabel('电流')
plt.plot(X,data[1],linewidth = '0.6')
plt.subplot(333)
plt.ylim((-10, 80))
plt.xticks(rotation = 20,fontsize=6)
plt.ylabel('前轴承温度3')
plt.plot(X,data[2],linewidth = '0.6')
plt.subplot(334)
plt.ylim((-10, 80))
plt.xticks(rotation = 20,fontsize=6)
plt.ylabel('中轴承温度3')
plt.plot(X,data[3],linewidth = '0.6')
plt.subplot(335)
plt.ylim((0, 200))
plt.xticks(rotation = 20,fontsize=6)
plt.ylabel('后轴承温度3')
plt.plot(X,data[4],linewidth = '0.6')
plt.subplot(336)
plt.ylim((-10, 80))
plt.xticks(rotation = 20,fontsize=6)
plt.ylabel('前轴承温度')
plt.plot(X,data[5],linewidth = '0.6')
plt.subplot(337)
plt.ylim((-10, 80))
plt.xticks(rotation = 20,fontsize=6)
plt.ylabel('后轴承温度')
plt.plot(X,data[6],linewidth = '0.6')
plt.subplot(338)
plt.ylim((-5, 15))
plt.xticks(rotation = 20,fontsize=6)
plt.ylabel('轴承振动方向X')
plt.plot(X,data[7],linewidth = '0.6')
plt.subplot(339)
plt.ylim((-5, 15))
plt.xticks(rotation = 20,fontsize=6)
plt.ylabel('轴承振动方向Y')
plt.plot(X,data[8],linewidth = '0.6')
plt.suptitle('一次风机历史数据清洗图')
plt.show()



# fig = plt.figure()
# plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
# plt.subplot(331)

# plt.xticks(rotation = 20,fontsize=6)
# plt.ylabel('调节挡板阀位')
# plt.plot(X,data[0],linewidth = '0.6')
# plt.subplot(332)

# plt.xticks(rotation = 20,fontsize=6)
# plt.ylabel('电流')
# plt.plot(X,data[1],linewidth = '0.6')
# plt.subplot(333)

# plt.xticks(rotation = 20,fontsize=6)
# plt.ylabel('前轴承温度3')
# plt.plot(X,data[2],linewidth = '0.6')
# plt.subplot(334)

# plt.xticks(rotation = 20,fontsize=6)
# plt.ylabel('中轴承温度3')
# plt.plot(X,data[3],linewidth = '0.6')
# plt.subplot(335)

# plt.xticks(rotation = 20,fontsize=6)
# plt.ylabel('后轴承温度3')
# plt.plot(X,data[4],linewidth = '0.6')
# plt.subplot(336)

# plt.xticks(rotation = 20,fontsize=6)
# plt.ylabel('前轴承温度')
# plt.plot(X,data[5],linewidth = '0.6')
# plt.subplot(337)

# plt.xticks(rotation = 20,fontsize=6)
# plt.ylabel('后轴承温度')
# plt.plot(X,data[6],linewidth = '0.6')
# plt.subplot(338)

# plt.xticks(rotation = 20,fontsize=6)
# plt.ylabel('轴承振动方向X')
# plt.plot(X,data[7],linewidth = '0.6')
# plt.subplot(339)

# plt.xticks(rotation = 20,fontsize=6)
# plt.ylabel('轴承振动方向Y')
# plt.plot(X,data[8],linewidth = '0.6')
# plt.suptitle('一次风机历史数据清洗图')
# plt.show()