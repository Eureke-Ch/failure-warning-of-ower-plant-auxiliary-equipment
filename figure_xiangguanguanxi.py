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
        data1.append(sheet_name.cell_value(i,6))
        data2.append(sheet_name.cell_value(i,8))

data1 = [];data2 = [];X = []
data_path1 = './data/14A17Process.xlsx'
#data_path2 = './data/14A17-7-12.xlsx'
read_xsls(data_path1)
#read_xsls(data_path2)
font1 = {
'weight' : 'normal',
'size'   : 12,
}
pccs = np.corrcoef(data1,data2)
print(pccs)
plt.xlabel('一次风机电动机前轴承温度/℃',font1)
plt.ylabel('一次风机X方向上的轴承振动/(mm/s)',font1)
mpl.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus']=False
plt.scatter(data1,data2,c = 'cyan',s = 10)

plt.show()

