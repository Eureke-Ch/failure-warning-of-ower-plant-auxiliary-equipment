from sklearn.cluster import KMeans
import numpy as np
import pylab as pl
from operator import itemgetter
from collections import OrderedDict,Counter
import xlrd
import matplotlib.pyplot as plt
def read_xsls(xlsx_path):
    data_xsls = xlrd.open_workbook(xlsx_path) 
    sheet_name = data_xsls.sheets()[0]  
    count_nrows = sheet_name.nrows  
    for i in range(1,count_nrows):
        a = []
        for j in range(9):
            a.append(sheet_name.cell_value(i,j+1))
        pointsall.append(a)

pointsall = []
data_xsls = 'clusting.xlsx'
read_xsls(data_xsls)
kmeans = KMeans(n_clusters=2, random_state=0).fit(pointsall)
Y = kmeans.predict(pointsall)

plt.plot(list(range(len(pointsall))),Y)
plt.show()
#print(Y[0:100])