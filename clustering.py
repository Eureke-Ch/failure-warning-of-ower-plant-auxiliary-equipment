import pylab as pl
from operator import itemgetter
from collections import OrderedDict,Counter
import xlrd

def read_xsls(xlsx_path):
    data_xsls = xlrd.open_workbook(xlsx_path) 
    sheet_name = data_xsls.sheets()[0]  
    count_nrows = sheet_name.nrows  
    for i in range(1,count_nrows,100):
        for j in range(9):
            pointsall[j].append(sheet_name.cell_value(i,j+1))

pointsall = [[]for i in range(9)]
data_xsls = 'clusting.xlsx'
read_xsls(data_xsls)
points = pointsall[0]#[1.1,1.2,1.3,2.1,2,3,5,6]
#print(points)
# 初始时每个点指派为单独一簇
groups = [idx for idx in range(len(points))]
# 计算每个点对之间的距离
disP2P = {}
for idx1,point1 in enumerate(points):
  for idx2,point2 in enumerate(points):
    if (idx1 < idx2):
      distance = pow(abs(point1-point2),2)
      disP2P[str(idx1)+"#"+str(idx2)] = distance
# 按距离降序将各个点对排序
disP2P = OrderedDict(sorted(disP2P.items(), key=itemgetter(1), reverse=True))
# 当前有的簇个数
groupNum = len(groups)
# 过分合并会带入噪音点的影响，当簇数减为finalGroupNum时，停止合并
finalGroupNum = 4
while groupNum > finalGroupNum:
  # 选取下一个距离最近的点对
  twopoins,distance = disP2P.popitem()
  pointA = int(twopoins.split('#')[0])
  pointB = int(twopoins.split('#')[1])
  pointAGroup = groups[pointA]
  pointBGroup = groups[pointB]
  # 当前距离最近两点若不在同一簇中，将点B所在的簇中的所有点合并到点A所在的簇中，此时当前簇数减1
  if(pointAGroup != pointBGroup):
    for idx in range(len(groups)):
      if groups[idx] == pointBGroup:
        groups[idx] = pointAGroup
    groupNum -= 1
# 选取规模最大的3个簇，其他簇归为噪音点
wantGroupNum = 3
finalGroup = Counter(groups).most_common(wantGroupNum)
finalGroup = [onecount[0] for onecount in finalGroup]
dropPoints = [points[idx] for idx in range(len(points)) if groups[idx] not in finalGroup]
# 打印规模最大的3个簇中的点
group1 = [points[idx] for idx in range(len(points)) if groups[idx]==finalGroup[0]]
group2 = [points[idx] for idx in range(len(points)) if groups[idx]==finalGroup[1]]
group3 = [points[idx] for idx in range(len(points)) if groups[idx]==finalGroup[2]]
pl.plot([eachpoint for eachpoint in group1], [0 for i in range(len(group1))], 'or')
pl.plot([eachpoint for eachpoint in group2], [0 for i in range(len(group2))], 'oy')
pl.plot([eachpoint for eachpoint in group3], [0 for i in range(len(group3))], 'og')
# 打印噪音点，黑色
pl.plot([eachpoint for eachpoint in dropPoints], [0 for i in range(len(dropPoints))], 'ok')
pl.show()






















#二维数据凝聚层次聚类
'''# scoding=utf-8
# Agglomerative Hierarchical Clustering(AHC)
import pylab as pl
from operator import itemgetter
from collections import OrderedDict,Counter
points = [[int(eachpoint.split('#')[0]), int(eachpoint.split('#')[1])] for eachpoint in open("points","r")]
print(points)
# 初始时每个点指派为单独一簇
groups = [idx for idx in range(len(points))]
# 计算每个点对之间的距离
disP2P = {}
for idx1,point1 in enumerate(points):
  for idx2,point2 in enumerate(points):
    if (idx1 < idx2):
      distance = pow(abs(point1[0]-point2[0]),2) + pow(abs(point1[1]-point2[1]),2)
      disP2P[str(idx1)+"#"+str(idx2)] = distance
# 按距离降序将各个点对排序
disP2P = OrderedDict(sorted(disP2P.items(), key=itemgetter(1), reverse=True))
# 当前有的簇个数
groupNum = len(groups)
# 过分合并会带入噪音点的影响，当簇数减为finalGroupNum时，停止合并
finalGroupNum = int(groupNum*0.5)
while groupNum > finalGroupNum:
  # 选取下一个距离最近的点对
  twopoins,distance = disP2P.popitem()
  pointA = int(twopoins.split('#')[0])
  pointB = int(twopoins.split('#')[1])
  pointAGroup = groups[pointA]
  pointBGroup = groups[pointB]
  # 当前距离最近两点若不在同一簇中，将点B所在的簇中的所有点合并到点A所在的簇中，此时当前簇数减1
  if(pointAGroup != pointBGroup):
    for idx in range(len(groups)):
      if groups[idx] == pointBGroup:
        groups[idx] = pointAGroup
    groupNum -= 1
# 选取规模最大的3个簇，其他簇归为噪音点
wantGroupNum = 3
finalGroup = Counter(groups).most_common(wantGroupNum)
finalGroup = [onecount[0] for onecount in finalGroup]
dropPoints = [points[idx] for idx in range(len(points)) if groups[idx] not in finalGroup]
# 打印规模最大的3个簇中的点
group1 = [points[idx] for idx in range(len(points)) if groups[idx]==finalGroup[0]]
group2 = [points[idx] for idx in range(len(points)) if groups[idx]==finalGroup[1]]
group3 = [points[idx] for idx in range(len(points)) if groups[idx]==finalGroup[2]]
pl.plot([eachpoint[0] for eachpoint in group1], [eachpoint[1] for eachpoint in group1], 'or')
pl.plot([eachpoint[0] for eachpoint in group2], [eachpoint[1] for eachpoint in group2], 'oy')
pl.plot([eachpoint[0] for eachpoint in group3], [eachpoint[1] for eachpoint in group3], 'og')
# 打印噪音点，黑色
pl.plot([eachpoint[0] for eachpoint in dropPoints], [eachpoint[1] for eachpoint in dropPoints], 'ok')
pl.show()
'''