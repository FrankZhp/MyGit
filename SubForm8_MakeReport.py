#! -*- coding utf-8 -*-
#! @Time  :2019/3/6 9:00
#! Author :Frank Zhang
#! @File  :SubForm_ComputePrice.py
#! Python Version 3.7

#三个文件：1、已经合并汇总好的Excel文件(4Summary) 2、当期价格表 3、待生成报表模板样例文件
#模块功能：计算价格


from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
from tkinter import ttk
import os
import time
import xlrd
import xlwt
from xlutils.copy import copy
import csv
import numpy as np
from collections import Counter



#模块功能：按照地区别、季度的价格，计算Ticket的价格
#第1个文件，已经汇总Mapping好的Excel文件，来自于步骤3
#第2个文件，价格表
#第3个文件，待生成报表的模板文件

"""
【处理逻辑】：
 1、读取Mapping File，按照Country和Period过滤一遍，对符合条件的计算费用，按照模板样例生成一个新的中间文件
 2、再进一步处理，按月（Price Table）和Support Type分Sheet
 

"""
def main():
    def selectSummaryFile():
        root.wm_attributes('-topmost',0)
        sfname = filedialog.askopenfilename(title='Please Select Summary File', filetypes=[('Excel', '*.xls'), ('All Files', '*')])
        #print(sfname)
        txt_MappingFile.insert(INSERT,sfname)
        root.wm_attributes('-topmost',1)

    def selectPriceFile():
        root.wm_attributes('-topmost',0)
        sfname = filedialog.askopenfilename(title='Please Select Price File', filetypes=[('xls', '*.xls'), ('All Files', '*')])
        #print(sfname)
        txt_PriceFile.insert(INSERT,sfname)
        root.wm_attributes('-topmost',1)

    def closeThisWindow():
        root.destroy()
        

    def doProcess():
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示','开始处理......')
        root.wm_attributes('-topmost',1)
        dictArea={0:'AU',1:'NZ',2:'CN',3:'HK',4:'TW',5:'SG',6:'TH',7:'ID',8:'MY',9:'VN',10:'PH'}
        dictPeriod={0:'2018Q11',1:'2018Q2',2:'2018Q3',3:'2018Q4',4:'2019Q1'}

        sMappingFile=txt_MappingFile.get()
        sPriceFile=txt_PriceFile.get()
        sModFile=os.getcwd()+"\\Mod\\FESTO_Invoice_Summary_Style2_Mod.xls"

        print(sModFile)
        wb1 = xlrd.open_workbook(filename=sMappingFile)
        wb2 = xlrd.open_workbook(filename=sPriceFile)
        
        wb3 = xlrd.open_workbook(filename=sModFile,formatting_info=True)                                 

        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows                                               #Summary File max row's number
        ncols1 = sheet1.ncols                                               #Summary File max col's number

        sheet2 = wb2.sheet_by_index(0)   
        nrows2 = sheet2.nrows                                               #Summary File max row's number
        ncols2 = sheet2.ncols                                               #Summary File max col's number
        rows = sheet2.row_values(0)                                         #获取行内容,第1行，列名
        cols = sheet2.col_values(1)                                         #获取列内容,第2列，SI Code
        
        wb = copy(wb3)
        ws = wb.get_sheet(1)
 
        sArea1=dictArea.get(lst_Area.current())
        #判断价格所在列
        if sArea1 in rows:
            iPriceCol=rows.index(sArea1)
            #ws.write(0,48,sArea1)

        sPeriod1=dictPeriod.get(lst_Period.current())
        
        iRow2=0
        lst_SICode=[]
        dict_SICode={}
        
        for iRow in range(1,nrows1):
            sTicketNo=sheet1.cell_value(iRow,0)
            sArea2=sheet1.cell_value(iRow,7).strip()
            #print('行号:'+str(iRow)+'Area:'+sArea2)
            sClosedDate=sheet1.cell_value(iRow,32).strip()
            sPeriod2=getPeriod(sClosedDate)
            sPriceTable=getPriceTable(sClosedDate)
            
            if sTicketNo!="" and sArea1==sArea2 and sPeriod1==sPeriod2:
                iRow2=iRow2+1

                for iCol in range(15):
                    ws.write(iRow2,iCol,sheet1.cell_value(iRow,iCol))
                
                for iCol in range(26,ncols1):
                    ws.write(iRow2,iCol,sheet1.cell_value(iRow,iCol))

                    
                #计算价格
                sServiceType=sheet1.cell_value(iRow,1).strip()                            #Service Type         B列
                sServiceSubType=sheet1.cell_value(iRow,2).strip()                         #Service Sub-Type     C列
                sServiceLevel=sheet1.cell_value(iRow,4).strip()                           #Service Level        E列 
                sTicketType=sheet1.cell_value(iRow,3).strip()                             #Ticket Type          D列 

                sSICode=sServiceType+"_"+sServiceSubType+"_"+sServiceLevel+"_"+sTicketType
                #print('SICode：'+sSICode)
                lst_SICode.append(sSICode)
                

                
                if sSICode in cols:
                    i=cols.index(sSICode)
                    if i>=0:
                        tempRows=sheet2.row_values(i)
                        #print("行："+str(i))
                        #print(tempRows)
                        sServiceItemNo=tempRows[0]
                        priceItem=tempRows[iPriceCol]
                        ws.write(iRow2,52,sServiceItemNo)                                   #BA列
                        ws.write(iRow2,53,sSICode)                                          #BB列
                        ws.write(iRow2,54,priceItem)                                        #BC列

                        #增加Period列和Price Table列（月）
                        ws.write(iRow2,55,sPeriod2)
                        ws.write(iRow2,56,sPriceTable) 
                else:
                    print('Not in Cols')
            else:
               # print('Row'+str(iRow)+'TicketNo:'+sTicketNo +'Area:'+sArea2)
                pass   

        #记录不同SI Code的Ticket的条数
        print('记录数：'+str(len(lst_SICode)))
        #print(all_np(lst_SICode))
        dict = {}
        for key in lst_SICode:
            dict[key] = dict.get(key, 0) + 1
        #print(dict)
        iRow2=iRow2+3
        ws.write(iRow2,0,"Service Item No")
        ws.write(iRow2,1,"SI Code")
        ws.write(iRow2,2,"Unit Price")
        ws.write(iRow2,3,"Ticket")
        ws.write(iRow2,4,"Total")
        totalPrice=0
        for k,v in dict.items():
            #print(k,v)
            if k in cols:
                sSICode=k
                i=cols.index(sSICode)
                if i>=0:
                    iRow2=iRow2+1
                    tempRows=sheet2.row_values(i)
                    #print("行："+str(i))
                    #print(tempRows)
                    sServiceItemNo=tempRows[0]
                    priceItem=tempRows[iPriceCol]
                    ws.write(iRow2,0,sServiceItemNo)
                    ws.write(iRow2,1,k)
                    ws.write(iRow2,2,priceItem)
                    ws.write(iRow2,3,v)
                    subtotal=int(v)*float(priceItem)
                    totalPrice=totalPrice+subtotal
                    ws.write(iRow2,4,subtotal)
                else:
                    print('Not in Cols')
        iRow2=iRow2+1
        ws.write(iRow2,0,"Total")
        ws.write(iRow2,4,totalPrice)
              
        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="8.ComputePriceSyeStyle2_"+strTime+".xls"
        wb.save(sPath+sReportName)        
       

        print("Work is over.")
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示','处理完毕，生成文件：\n'+sPath+sReportName)   
        root.wm_attributes('-topmost',1)

    def getPriceTable(sClosedDate):
        sPriceTable=""
        sYear=""
        sMonth=""
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


    
    #初始化
    root=Tk()

    #设置窗体标题
    root.title('Compute Price---Make Report-----Step 8 ----Style 2----Single sheet')

    #设置窗口大小和位置
    root.geometry('660x300+430+220')


    lbl_Summary=Label(root,text='Mapping File:',width=20,justify = tkinter.RIGHT,)
    txt_MappingFile=Entry(root,bg='white',width=66)
    btn_browse1=Button(root,text='Browse',width=8,command=selectSummaryFile)

    lbl_Prices=Label(root,text='Price File:',width=16,justify = tkinter.RIGHT)
    txt_PriceFile=Entry(root,bg='white',width=66)
    btn_browse2=Button(root,text='Browse',width=8,command=selectPriceFile)

    lbl_Area=Label(root,text="Select Country:",width=16,justify = tkinter.RIGHT)

    sArea = tkinter.StringVar()
    lst_Area=ttk.Combobox(root,width=20,textvariable=sArea) #下拉列表框
    lst_Area['values']=('AU','NZ','CN','HK','TW','SG','TH','ID','MY','VN','PH')
 

    lbl_Quarter=Label(root,text="Select Period:",justify = tkinter.RIGHT)

    #r = tkinter.StringVar()
    #r.set("Q1")
    
    #rdo_Quarter1 = tkinter.Radiobutton(root,
    #                                  variable = r,
    #                                  value = "Q1",
    #                                  text = "Q1")
    #rdo_Quarter2 = tkinter.Radiobutton(root,
    #                                  variable = r,
    #                                  value = "Q2",
    #                                  text = "Q2")
    #rdo_Quarter3 = tkinter.Radiobutton(root,
    #                                  variable = r,
    #                                  value = "Q3",
    #                                  text = "Q3")
    #rdo_Quarter4 = tkinter.Radiobutton(root,
    #                                  variable = r,
    #                                  value = "Q4",
    #                                  text = "Q4")
    sPeriod = tkinter.StringVar()
    lst_Period=ttk.Combobox(root,width=20,textvariable=sPeriod) #下拉列表框
    lst_Period['values']=('2018Q1','2018Q2','2018Q3','2018Q4','2019Q1')

    
    btn_process=Button(root,text='Process',width=8,command=doProcess)
    btn_exit=Button(root,text='Exit',width=8,command=closeThisWindow)
 

    lbl_Summary.pack()
    txt_MappingFile.pack()
    btn_browse1.pack()

    lbl_Prices.pack()
    txt_PriceFile.pack()
    btn_browse2.pack()



    lbl_Area.pack()

    lbl_Quarter.pack()
    #rdo_Quarter1.pack()
    #rdo_Quarter2.pack()
    #rdo_Quarter3.pack()
    #rdo_Quarter4.pack()
    lst_Period.pack()
    
    btn_process.pack()
    btn_exit.pack() 

    lbl_Summary.place(x=26,y=30)
    txt_MappingFile.place(x=139,y=30)
    btn_browse1.place(x=550,y=26)

    lbl_Prices.place(x=50,y=60)
    txt_PriceFile.place(x=139,y=60)
    btn_browse2.place(x=550,y=56)

    lbl_Area.place(x=35,y=90)
    lst_Area.place(x=139,y=90)

    lbl_Quarter.place(x=300,y=90)
    #rdo_Quarter1.place(x=390,y=90)
    #rdo_Quarter2.place(x=430,y=90)
    #rdo_Quarter3.place(x=470,y=90)
    #rdo_Quarter4.place(x=510,y=90)
    lst_Period.place(x=385,y=90)
    
    btn_process.place(x=250,y=160)
    btn_exit.place(x=350,y=160)
 
 
    root.mainloop() 

 
if __name__=="__main__":
    main()
