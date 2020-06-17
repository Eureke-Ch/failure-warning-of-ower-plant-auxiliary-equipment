import xlrd
from sklearn.svm import OneClassSVM
import matplotlib.pyplot as plt
def read_xsls(xlsx_path):
    data_xsls = xlrd.open_workbook(xlsx_path) 
    sheet_name = data_xsls.sheets()[0]  
    count_nrows = sheet_name.nrows 
    for i in range(1,count_nrows,40):
        if i < 3000 or (i>40000 and i<100000) or (i > 240000 and i< 270000) or (i>320000 and i<360000) or (i>840000 and i<990000):
            pass
        else:
            a = []
            for j in range(9):
                a.append(sheet_name.cell_value(i,j+1))
            data.append(a)
          
#读取数据
data_path1 = './data/13A17ProAll.xlsx'
data = []
read_xsls(data_path1) 

clf = OneClassSVM(gamma='auto',nu = 0.001).fit(data)


Y = []
data_xsls = xlrd.open_workbook("./data/13A17ProAll.xlsx") 
sheet_name = data_xsls.sheets()[0]  
#count_nrows = sheet_name.nrows  
for i in range(30000,60000):
    a = []
    for j in range(9):
        a.append(sheet_name.cell_value(i,j+1))
    Y.append(clf.predict([a]))

Z = []
for i in range(len(Y)):
    n = 0
    if i < 19:
        Z.append(Y[i])
    else:
        for j in range(20):
            n += (Y[i-j]/20)
        Z.append(n)
X = []
for i in range(len(Y)):
    X.append(i+30000)
plt.plot(X,Z)
plt.show()