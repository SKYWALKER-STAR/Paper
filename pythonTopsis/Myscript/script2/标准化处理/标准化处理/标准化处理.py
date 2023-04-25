import pandas as pd
from openpyxl import load_workbook;
import numpy as np
import secondStep

#该函数用来计算熵权中的标准化步骤
def dataStanderlization(file_path:str,columnName:list,columnWriteName:str):
    df = pd.read_excel(file_path)
    for n in range(0,15):
        value = df.loc[n,columnName]
        if value > 0:
            value_computed = (df.loc[n,columnName] - df.loc[17,columnName])/df.loc[18,columnName]
        elif value < 0:
            value_computed = (df.loc[16,columnName] - df.loc[n,columnName]) / df.loc[18,columnName]
        
        df.loc[n,columnWriteName] = value_computed;
        df.to_excel(file_path,index=False)

def sumOfStdResult(filePath:str):
    sum = 0;
    df = pd.read_excel(filePath)
    for i in range(0,len(df)):
        sum += df.loc[i,"result_std_2021"]
    
    df.loc[19,"result_std_2021"] = sum;
    df.to_excel(filePath)


def computePij(filePath:str):
    df = pd.read_excel(filePath)
    sumOfBij = df.loc[19,"result_std_2021"]

    sumOfPij = 0
    for i in range(0,len(df)):
        singleBij = df.loc[i,"result_std_2021"]
        Pij = (singleBij/sumOfBij);

        df.loc[i,"P(ij)_2021"] = Pij;

        if Pij != 0:
            df.loc[i,"LnP(ij)_2021"] = np.log(Pij)
        sumOfPij += df.loc[i,"P(ij)_2021"]

    df.loc[20,"P(ij)_2021"] = sumOfPij
    df.to_excel(filePath,index=False)

#这个函数计算熵值，熵值的计算是以每一个企业的每一个指标进行计算，也就是说，该函数要跨表完成任务。
def computeTheFuckingEntropy(pathroot:str,sheet_list:list):

    company_list = []

    #该循环打开所有的表，并且存储再company_list中
    for i in sheet_list:
        filePath = pathroot+i+".xlsx"
        df = pd.read_excel(filePath)
        company_list.append(df) #该列表存储pd对象


    list_of_index = [];#该列表用来记录每个指标的信息熵
    list_of_index_with_ln = [];#该列表用来记录信息熵除以lnm后的值
    list_of_differantial_value = []#记录变异系数
    list_of_weight = [] #记录熵权
    res = 0;
    #该循环记录下每一个Pij与Lnpij的乘积的和，存储在一个list中
    for j in range(0,15):
        res = 0;

        for i in company_list:
            value_pij = i.loc[j,"P(ij)_2021"]
            value_lnpij = i.loc[j,"LnP(ij)_2021"]

            if (np.isnan(value_pij)) != True and (np.isnan(value_lnpij)) != True:
                res += value_pij * value_lnpij
        list_of_index.append(-(res))


    for i in list_of_index:
        value_final = i/np.log(15)
        list_of_index_with_ln.append(value_final)

    for i in range(0,len(company_list)):
        pathfile=pathroot+sheet_list[i]+".xlsx"
        for j in range(0,len(list_of_index_with_ln)):
            company_list[i].loc[j,"entropy_2021"] = list_of_index_with_ln[j]
        company_list[i].to_excel(pathfile,index=False)

    #计算指标的差异度
    sumOfDifferantialValue = 0;
    for i in list_of_index_with_ln:
        differantialValue_temp = 0;
        differantialValue_temp = 1-i;
        list_of_differantial_value.append(differantialValue_temp)
        sumOfDifferantialValue += differantialValue_temp;

    for i in list_of_differantial_value:
        weight_temp = i/sumOfDifferantialValue
        list_of_weight.append(weight_temp)

    for i in range(0,len(company_list)):
        pathfile=pathroot+sheet_list[i]+".xlsx"
        for j in range(0,len(list_of_weight)):
            company_list[i].loc[j,"W(ij)_2021"] = list_of_weight[j]
        company_list[i].to_excel(pathfile,index=False)

def computeRij(filePath:str):
    df = pd.read_excel(filePath)
    for i in range(0,15):
        originValue = df.loc[i,"value_2021"];
        weightValue = df.loc[i,"W(ij)_2021"];
        resultValue = originValue * weightValue;
        df.loc[i,"r(ij)_2021"] = resultValue;

    df.to_excel(filePath,index=False)

def insertValueToFile(filePath:str,value:float,positionRow:int,positionCol:str):
    df = pd.read_excel(filePath)
    df.loc[positionRow,positionCol] = value;

#该函数将以元和人为单位的数值转化为以亿为单位
def transformUnits(filePath:str):
    df = pd.read_excel(filePath)
    for i in range(0,15):
        value = df.loc[i,"value_2021"]
        if value > 1000:
            value = value / 100000000
            df.loc[i,"value_2021"] = value

    df.to_excel(filePath,index=False)

#增加max、min等数
def modifycation(filePath:str):
    df = pd.read_excel(filePath)
    df.loc[16,"评价指标"] = "max"
    df.loc[17,"评价指标"] = "min"
    df.loc[18,"评价指标"] = "max-min"
    df.loc[19,"评价指标"] = "SumOfResultStd"
    df.loc[20,"评价指标"] = "SumOfPij"

    value_list = []
    for i in range(0,15):
        value = df.loc[i,"value_2021"]
        value_list.append(value)

    max_value = max(value_list)
    min_value = min(value_list)

    maxMinusmin = max_value - min_value;

    df.loc[16,"value_2021"] = max_value;
    df.loc[17,"value_2021"] = min_value;
    df.loc[18,"value_2021"] = maxMinusmin
    df.to_excel(filePath,index=False)

def modifyNo(filePath:str):
    df = pd.read_excel(filePath)
    for i in range(0,len(df)):
        df.loc[i,"序号"] = i;

    df.to_excel(filePath,index=False)

#该函数用来展示表的信息
def showInfo(file_path:str,sheet_name:str):
    ds = pd.read_excel(file_path,sheet_name)
    print(ds.info())

#该函数用来将收集到的数据转移到用来操作的电子表格中
def migration(pathroot:str,path_collect:str,sheet_list:list):
        for i in sheet_list:
            ds = pd.read_excel(path_collect,i)
            ds.to_excel(pathroot+i+".xlsx",index=False)

#该函数用来插入相应的列
def doTheInsert(pathroot:str,sheet_list:list):
    for i in sheet_list:
       ds = pd.read_excel(pathroot+i+".xlsx")
       ds.insert(loc=4,column="result_std_2021",value=0)
       ds.insert(loc=5,column="P(ij)_2021",value=0)
       ds.insert(loc=6,column="LnP(ij)_2021",value=0)
       ds.insert(loc=7,column="entropy_2021",value=0)
       ds.insert(loc=8,column="W(ij)_2021",value=0)
       ds.insert(loc=9,column="r(ij)_2021",value=0)

       ds.insert(loc=11,column="result_std_2020",value=0)
       ds.insert(loc=12,column="P(ij)_2020",value=0)
       ds.insert(loc=13,column="LnP(ij)_2020",value=0)
       ds.insert(loc=14,column="entroy_2020",value=0)
       ds.insert(loc=15,column="W(ij)_2020",value=0)
       ds.insert(loc=16,column="r(ij)_2020",value=0)
       ds.to_excel(pathroot+i+".xlsx",index=False)

#该函数用来修改表中的列和行的信息，该函数调用的函数都是一次处理一个表
def loopEverySheets(pathroot:str,sheet_list:list):
    for i in sheet_list:
        filePath = pathroot+i+".xlsx"
        transformUnits(filePath)
        modifycation(filePath)
        dataStanderlization(filePath,"value_2021","result_std_2021")
        modifyNo(filePath)
        sumOfStdResult(filePath)
        computePij(filePath)

def loopEverySheetsForRij(pathroot:str,sheet_list:list):
    for i in sheet_list:
        filePath = pathroot+i+".xlsx"
        computeRij(filePath)

#该函数是入口函数
def main():
    sheet_list = ["dahuagufen","haikangweishi","hengyuxintong","kedaxunfei","shiyuangufen","zhongxintongxun","dongfangdianzi","hangxinkeji","tcl","xinguodu"]#企业名称
    path_collect = "C:\\Users\\tu\\Desktop\\excelTest\\评价指标_数据收集.xlsx"
    path_compute = "C:\\Users\\tu\\Desktop\\excelTest\\评价指标_熵权计算.xlsx"
    pathroot = "C:\\Users\\tu\\Desktop\\excelTest\\熵权计算_表格们\\" #数据集所在的根目录

    #migration(pathroot,path_collect,sheet_list);
    #doTheInsert(pathroot,sheet_list);
    #loopEverySheets(pathroot,sheet_list);
    #computeTheFuckingEntropy(pathroot,sheet_list)
    #loopEverySheetsForRij(pathroot,sheet_list)

    m = secondStep.MaxMinVectors(pathroot,sheet_list);
    listToCompute = m.computeMaxMinVector_2021();
    print("Positive:")
    m.computeDijPositive()
    m.printList(m.D_Positive)
    print("Negtive")
    m.computeDijNegtive()
    m.printList(m.D_Negtive)

    print("Final")
    m.computeFinalResult()
if __name__ == "__main__":
    main()