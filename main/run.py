#coding=utf-8

import xlrd
import os
from xlutils.copy import copy
from ddt import ddt,data,unpack

class OperateExcel(object):

    def __init__(self,fileName):
        self.path = os.path.dirname(os.getcwd())+"\data\\"
        # print(self.path)
        self.fileName = fileName
        self.table = xlrd.open_workbook(self.path + self.fileName)
        self.sheet = self.table.sheet_by_index(0)

    #筛选字符串大于num的内容
    def filtration_value(self,col,num):
        self.dataList = self.sheet.col_values(col)
        # print(self.dataList)
        list = []
        for i in self.dataList:
            if type(i)== float:
                i = str(int(i))
                # print(type(i),i,len(i))
            if len(i)>= num:
                list.append(i)
        return list

    #将结果写到excel中
    def write_colvalue(self,row,col,value):
        table = xlrd.open_workbook(self.path + self.fileName)
        table_copy = copy(table)
        sheet_copy = table_copy.get_sheet(0)
        sheet_copy.write(row,col,value)
        table_copy.save(self.path + self.fileName)


    def get_cellvalue(self,col,set_num,write_col):
        #获取所有行数
        row_num = self.sheet.nrows
        print("行数：",row_num)
        for num in range(14033,row_num):
            cell = self.sheet.cell_value(num,col)
            print(num,cell)
            if type(cell) == float:
                cell = str(int(cell))
            if len(cell)>=set_num:
                self.write_colvalue(num,write_col,cell)



if __name__ == '__main__':

    Opexcele = OperateExcel("testdemo.xls")
    # list = Opexcele.filtration_value(6,5)
    # print(list)
    Opexcele.get_cellvalue(6,5,8)
