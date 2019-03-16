# -*- coding:utf-8 -*-
#模块功能：判断某个文件夹下有几个Excel文件，每个Excel有几个Sheet及Sheet Name

import os
import openpyxl
import xlrd



def getFileNames(path):
    filenames = os.listdir(path)
    for i, filename in enumerate(filenames):
         if i==0:
            iSpecialFile=i+1
            sFileName=filename

         print('==================第%s个文件========================='%(i+1))
         print('文件名：%s'%(filename))
         getSheetNames(path,filename)
    print('\n')
    print('--------------------选择指定的第几个文件-------------------------')
    print('指定的是第%s个文件:'%iSpecialFile+sFileName )
    print('----------------------------------------------------------------')

def getSheetNames(path,sFileName):
    wb1 = openpyxl.load_workbook(path+'\\'+sFileName)
    wb2 = xlrd.open_workbook(filename=path+'\\'+sFileName)
    # 获取workbook中所有的表格
    sheets = wb1.sheetnames

    sheet1=wb2.sheet_by_index(0)
    rows=sheet1.row_values(2)
    print(rows)
    # 循环遍历所有sheet
    for i in range(len(sheets)):
        sheet = wb1[sheets[i]]
        print('第' + str(i + 1) + '个sheet Name: ' + sheet.title)
    

if __name__=='__main__':
    path=r'C:\\Work\\Python\\MergeExcel\\Source'
    getFileNames(path)
