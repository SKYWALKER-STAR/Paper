"""
该文件中的函数计算正负理想解与欧氏距离
"""

import pandas as pd
import numpy as np
import math
class MaxMinVectors():
    company_list = []
    pd_list = []
    MaxVector = []
    MinVector = []
    pathroot:str

    D_Positive = [] #存储每个企业到最优解的距离
    D_Negtive = [] #存储每个企业到最劣解的距离

    def __init__(self,root:str,companies:list):
        self.patthroot = root
        self.company_list = companies
        for i in self.company_list:
            pd_temp = pd.read_excel(root+i+".xlsx")
            self.pd_list.append(pd_temp)

    def printPDinfos(self):
        for i in self.pd_list:
            print(i.info())

    def computeMaxMinVector_2021(self):
        for j in range(0,15):
            rij_list = []
            for i in self.pd_list:
                rij_list.append(i.loc[j,"r(ij)_2021"])
            self.MaxVector.append(max(rij_list))
            self.MinVector.append(min(rij_list))
        return 0

    def computeDijPositive(self):
        for i in self.pd_list:
            SumOfMaxMinusOriginPower = 0
            for j in range(0,15):
                MaxMinusOrigin = self.MaxVector[j] - i.loc[j,"r(ij)_2021"]
                if (np.isnan(MaxMinusOrigin)) != True and MaxMinusOrigin != 0:
                    MaxMinusOriginPower = math.pow(MaxMinusOrigin,2)
                    SumOfMaxMinusOriginPower += MaxMinusOriginPower 
            DijResultPositive = math.sqrt(SumOfMaxMinusOriginPower)
            self.D_Positive.append(DijResultPositive)

    def computeDijNegtive(self):
        for i in self.pd_list:
            SumOfMinMinusOriginPower = 0
            for j in range(0,15):
                MinMinusOrigin = self.MinVector[j] - i.loc[j,"r(ij)_2021"]
                if (np.isnan(MinMinusOrigin)) != True and MinMinusOrigin != 0:
                    MinMinusOriginPower = math.pow(MinMinusOrigin,2)
                    SumOfMinMinusOriginPower += MinMinusOriginPower
            DijResultNegtive = math.sqrt(SumOfMinMinusOriginPower)
            self.D_Negtive.append(DijResultNegtive)

    def computeFinalResult(self):
        if len(self.D_Negtive) != len(self.D_Positive):
            print("Error at computeFinalResult")
            exit(-1)
        for i in range(0,len(self.D_Negtive)):
            res = self.D_Negtive[i] / (self.D_Negtive[i] + self.D_Positive[i])
            print(res)
                                                                                                                            
    def printList(self,aList:list):
        for i in range(0,len(aList)):
            print(aList[i])

if __name__ == "__main__":
    print("Hello from main")