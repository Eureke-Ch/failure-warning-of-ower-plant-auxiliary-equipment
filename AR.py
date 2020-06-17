import xlrd
import xlwt
import matplotlib.pyplot as plt
import pylab as pl
from operator import itemgetter
from collections import OrderedDict,Counter
def read_xsls(xlsx_path):
    data_xsls = xlrd.open_workbook(xlsx_path) 
    sheet_name = data_xsls.sheets()[0]  
    count_nrows = sheet_name.nrows  
    for i in range(1,count_nrows):
        n = []
        for j in range(9):
            n.append(sheet_name.cell_value(i,j+1))
        data.append(n)
    
        

data_path = 'li_san.xlsx'
data = []
read_xsls(data_path)

def loadDataSet():
    return data#[[1,2,5],[2,4],[2,3],[1,2,4],[1,3],[2,3],[1,3],[1,2,3,5],[1,2,3]]

#寻找频繁项集
def createC1(dataSet):  #创造候选项集C1，C1是大小为1的所有候选项集的集合
    C1 = []
    for transaction in dataSet:   
        for item in transaction:
            if not [item] in C1:
                C1.append([item])
    #C1.sort()     
    return list(map(frozenset,C1))
 
def scanD(D,Ck,minSupport):  #此函数计算支持度,筛选满足要求的项集成为频繁项集Lk，D是数据集，Ck为候选项集C1或C2或C3 ...
    ssCnt = {}
    for tid in D:
        for can in Ck:
            if can.issubset(tid):
                if not can in ssCnt:
                    ssCnt[can] = 1
                else:
                    ssCnt[can] += 1
    numItems = float(len(D))
    retList = []
    supportData = {}
    for key in ssCnt:
        support = ssCnt[key]/numItems    # 计算支持度
        if support >= minSupport:   
           # retList.insert(0,key)
            retList.append(key)
        supportData[key] = support
    return retList, supportData
 
def aprioriGen(Lk, k):   
    lenLk = len(Lk)
    temp_dict = {}  
    for i in range(lenLk):
        for j in range(i+1, lenLk):
            L1 = Lk[i]|Lk[j]  
            if len(L1) == k: 
                if not L1 in temp_dict:  
                    temp_dict[L1] = 1
    return list(temp_dict)  


def apriori(dataSet, minSupport = 0.5):  # 通过循环得出[L1,L2,L3..]频繁项集列表
    C1 = createC1(dataSet)    
    D = list(map(set,dataSet))
    L1,supportData = scanD(D, C1, minSupport)  
    L = [L1]
    k = 2
    while (len(L[k-2]) > 0):   
        Ck = aprioriGen(L[k-2],k)
        Lk, supK = scanD(D,Ck,minSupport)
        supportData.update(supK)
        L.append(Lk)
        k += 1
    return L,supportData

 

def calcConf(freqSet, H, supportData, br1, minConf = 0.7): # 筛选符合可信度要求的规则，并返回符合可信度要求的右件
    prunedH = []  # 存储符合可信度的右件
    for conseq in H: 
        conf = supportData[freqSet]/supportData[freqSet-conseq] 
        if conf>= minConf:
            print(freqSet-conseq,"-->",conseq,"\tconf:",conf,file = f)
            br1.append((freqSet-conseq,conseq,conf))
            prunedH.append(conseq)
    return prunedH
 



#发现关联规则
def rulesFromConseq(freqSet, H, supportData, brl, minConf=0.7):
    m = len(H[0])
    if (len(freqSet) > (m + 1)): # 尝试进一步合并
        Hmp1 = aprioriGen(H, m+1) # 将单个集合元素两两合并
        Hmp1 = calcConf(freqSet, Hmp1, supportData, brl, minConf)
        if (len(Hmp1) > 1):    
            rulesFromConseq(freqSet, Hmp1, supportData, brl, minConf)

 
def generateRules(L, supportData, minConf=0.7):  # 产生规则
    bigRuleList = []
    for i in range(1, len(L)):  
        for freqSet in L[i]:
            H1 = [frozenset([item]) for item in freqSet]
            if i>1:  
                rulesFromConseq(freqSet, H1, supportData, bigRuleList, minConf)
            else:  
                calcConf(freqSet, H1, supportData, bigRuleList, minConf)
    return bigRuleList


f = open("Rules.txt", "w") 
dataset=loadDataSet()
L,supportData=apriori(dataset,minSupport = 0.01)
print(L,supportData, file = f)
y = 0
for i in range(len(L)):
    y+=len(L[i])
print(y)
rules=generateRules(L,supportData,minConf=0.95)


path = 'Rules.xls'
rb = xlwt.Workbook()
sheet = rb.add_sheet('sheet1')
for i in range(len(rules)):
    mm = 0;nn = 0
    for j in rules[i][0]:
        sheet.write(i+1,mm+1,j)
        mm +=1
    for j in rules[i][1]:
        sheet.write(i+1,nn+11,j)
        nn +=1
rb.save(path)





