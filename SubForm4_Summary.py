from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
import xlrd
import xlwt
from xlutils.copy import copy
import os
import time
import datetime

#Menu4
#功能模块：从已Mapping的Excel File转换格式输出文件

def main():
    def selectExcelfile1():
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '*.xls'), ('All Files', '*')])
        txt_MappingFile.insert(INSERT,sfname)

    def closeThisWindow():
        root.destroy()

    #第一个文件是3.Mapping Excel
    #第二个文件是Summary_Mod.xls
    def doProcess():
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示','开始处理......')
        root.wm_attributes('-topmost',1)
        sMappingFile=txt_MappingFile.get()
        #sClosedFile=text2.get()
        sModFile=os.getcwd()+"\\Mod\\Summary_mod.xls"

        print(sModFile)
        wb1 = xlrd.open_workbook(filename=sMappingFile)
        wb2 = xlrd.open_workbook(filename=sModFile, formatting_info=True)
        #wb2 = xlrd.open_workbook(filename=sModFile)                                 #Summary Mod Excel

        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows

        sheet2 = wb2.sheet_by_index(0)   
        nrows2 = sheet2.nrows
        ncols2 = sheet2.ncols

        wb3 = copy(wb2)
        ws3 = wb3.get_sheet(0)

        #填写生成文件从AA列开始的台头
        #for i in range(ncols2):
        #     ws3.write(0,i+26,sheet2.cell_value(0,i))

        for i in range(1,nrows1):
            sTicketNo1=sheet1.cell_value(i,0)
            ws3.write(i,0,sheet1.cell_value(i,0))
            #TicketNo
            #B列Period,格式：yyyyQX,来自closedDate
            sClosedDate=sheet1.cell_value(i,32)
            sPeriod=getPeriod(sClosedDate)
            sPriceTable=getPriceTable(sClosedDate)
            ws3.write(i,1,sPeriod)
            ws3.write(i,9,sPriceTable)

            ws3.write(i,10,sheet1.cell_value(i,26))                       #Priority
            ws3.write(i,11,sheet1.cell_value(i,27))                       #Ticket No
            ws3.write(i,12,sheet1.cell_value(i,28))                       #  
            ws3.write(i,13,sheet1.cell_value(i,29))                        
            ws3.write(i,14,sheet1.cell_value(i,30))
            ws3.write(i,15,sheet1.cell_value(i,31))
            ws3.write(i,16,sheet1.cell_value(i,32))                       #closedDate
            ws3.write(i,17,sheet1.cell_value(i,33))                      
            ws3.write(i,18,sheet1.cell_value(i,34))
            ws3.write(i,19,sheet1.cell_value(i,35))
            ws3.write(i,20,sheet1.cell_value(i,36))
            ws3.write(i,21,sheet1.cell_value(i,37))                        
            ws3.write(i,22,sheet1.cell_value(i,38))
            ws3.write(i,23,sheet1.cell_value(i,39))
            ws3.write(i,24,sheet1.cell_value(i,40))
            ws3.write(i,25,sheet1.cell_value(i,41))                      #Country
             
            ws3.write(i,26,sheet1.cell_value(i,0))                       #TicketNo
            ws3.write(i,27,sheet1.cell_value(i,1))                       #Service Type						
            ws3.write(i,28,sheet1.cell_value(i,2))                       #Service Sub-Type
            ws3.write(i,29,sheet1.cell_value(i,3))                       #Ticket Type
            ws3.write(i,30,sheet1.cell_value(i,4))                       #Service Level
            ws3.write(i,31,sheet1.cell_value(i,5))                       #Work Order No
            ws3.write(i,32,sheet1.cell_value(i,6))                       #Suport type
            ws3.write(i,33,sheet1.cell_value(i,7))                       #Country
            ws3.write(i,34,sheet1.cell_value(i,8))                       #Site Location
            ws3.write(i,35,sheet1.cell_value(i,9))                       #AssetID   
            ws3.write(i,36,sheet1.cell_value(i,10))                      #User ID
            ws3.write(i,37,sheet1.cell_value(i,11))                      #Status
            ws3.write(i,38,sheet1.cell_value(i,12))                      #Partner Ticket solved onsite Date    
            ws3.write(i,39,sheet1.cell_value(i,13))                      #Support Engineer Name
            ws3.write(i,40,sheet1.cell_value(i,14))                      #Resolution
        

        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="4.Summary_"+strTime+".xls"
        wb3.save(sPath+sReportName)        
       

        print("Work is over.")
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示','处理完毕,：\n'+sPath+sReportName)
        root.wm_attributes('-topmost',1)

    def getPriceTable(sClosedDate):
        sPriceTable=""
        sYear=""
        sMonth=""
        #sDay=""
        if sClosedDate!="":
            list1=sClosedDate.split(" ")
            sDate=list1[0]
            list2=sDate.split("/")
            i=0
            for row in list2:
                if i==0:
                    sMonth=row
                #if i==1:
                #    sDay=row
                if i==2:
                    sYear=row
                i=i+1
            sPriceTable=sYear+"-"+sMonth
        return sPriceTable
    def getPeriod(sClosedDate):
        sPeriod=""
        sYear=""
        sMonth=""
        sDay=""
        try:
            if sClosedDate!="":
                list1=sClosedDate.split(" ")
                sDate=list1[0]
                list2=sDate.split("/")
                i=0
                for row in list2:
                    if i==0:
                        sMonth=row
                    if i==1:
                        sDay=row
                    if i==2:
                        sYear=row
                    i=i+1
            dict={1:'Q1',2:'Q1',3:'Q1',4:'Q2',5:'Q2',6:'Q2',7:'Q3',8:'Q3',9:'Q3',10:'Q4',11:'Q4',12:'Q4'}
            sPeriod=sYear+dict.get(int(sMonth))
        except:
            print(sClosedDate)
        return sPeriod 

    def is_valid_date(str):
        '''判断是否是一个有效的日期字符串'''
        try:
            time.strptime(str, "%Y-%m-%d")
            return True
        except:
            return False

    #初始化
    root=Tk()

    #设置窗体标题
    root.title('Summary ---Step 4')

    #设置窗口大小和位置
    root.geometry('660x300+430+220')


    label1=Label(root,text='Select Mapping File:')
    txt_MappingFile=Entry(root,bg='white',width=66)
    btn_browse=Button(root,text='Browse',width=8,command=selectExcelfile1)

    
    btn_process=Button(root,text='Process',width=8,command=doProcess)
    btn_exit=Button(root,text='Exit',width=8,command=closeThisWindow)
 

    label1.pack()
    txt_MappingFile.pack()
    btn_browse.pack()

    
    btn_process.pack()
    btn_exit.pack() 

    label1.place(x=30,y=30)
    txt_MappingFile.place(x=146,y=30)
    btn_browse.place(x=550,y=26)

    
    btn_process.place(x=260,y=100)
    btn_exit.place(x=360,y=100)
 
 
    root.mainloop() 

 
if __name__=="__main__":
    main()
