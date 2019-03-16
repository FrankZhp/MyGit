#! -*- coding utf-8 -*-
#! @Time  :2019/3/6 9:00
#! Author :Frank Zhang
#! @File  :SubForm_ComputePrice.py
#! Python Version 3.7

#三个文件：1、已经合并汇总好的Excel文件 2、当期价格表 3、待生成报表模板样例文件
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


#模块功能：按照地区别、季度的价格，计算Ticket的价格
#第1个文件，已经汇总Mapping好的Excel文件
#第2个文件，价格表
#第3个文件，待生成报表的模板文件

def main():
    def selectSummaryFile():
        sfname = filedialog.askopenfilename(title='Please Select Summary File', filetypes=[('Excel', '*.xls'), ('All Files', '*')])
        #print(sfname)
        txt_SummaryFile.insert(INSERT,sfname)

    def selectPriceFile():
        sfname = filedialog.askopenfilename(title='Please Select Price File', filetypes=[('xls', '*.xls'), ('All Files', '*')])
        #print(sfname)
        txt_PriceFile.insert(INSERT,sfname)

    def closeThisWindow():
        root.destroy()

    def doProcess():
        tkinter.messagebox.showinfo('提示','开始处理')
        #print(r.get())
        dict={0:'AU',1:'NZ',2:'CN',3:'HK',4:'TW',5:'SG',6:'TH',7:'ID',8:'MY',9:'VN',10:'PH'}

        sSummaryFile=txt_SummaryFile.get()
        sPriceFile=txt_PriceFile.get()
        sModFile=os.getcwd()+"\\Mod\\FESTO Summary via USU & CS XXX_Mod.xls"

        print(sModFile)
        wb1 = xlrd.open_workbook(filename=sSummaryFile)
        wb3 = xlrd.open_workbook(filename=sPriceFile)
        #csv_file = csv.reader(open(sPriceFile))
        #print(csv_file)
        
        wb2 = xlrd.open_workbook(filename=sModFile,formatting_info=True)                                 

        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows                                               #Summary File max row's number
        ncols1 = sheet1.ncols                                               #Summary File max col's number

      

        sheet3 = wb3.sheet_by_index(0)   
        nrows3 = sheet3.nrows                                               #Summary File max row's number
        ncols3 = sheet3.ncols                                               #Summary File max col's number
        rows = sheet3.row_values(0)                                         #获取行内容,第1行，列名
        cols = sheet3.col_values(1)                                         #获取列内容,第2列，SI Code
        
    
        #print(rows)
        #print(cols)
       
     
        wb = copy(wb2)
        ws = wb.get_sheet(1)

        sArea1=dict.get(lst_Area.current())
        #判断价格所在列
        if sArea1 in rows:
            iPriceCol=rows.index(sArea1)
            ws.write(0,48,sArea1)


        iRow2=0
        for iRow in range(1,nrows1):
            sTicketNo=sheet1.cell_value(iRow,0)
            sArea2=sheet1.cell_value(iRow,33).strip()
            print('行号:'+str(iRow)+'Area:'+sArea2)
            if sTicketNo!="" and sArea1==sArea2:
                iRow2=iRow2+1
                for iCol in range(ncols1):
                    ws.write(iRow2,iCol,sheet1.cell_value(iRow,iCol))

                    
                #计算价格
                sServiceType=sheet1.cell_value(iRow,27).strip()
                sServiceSubType=sheet1.cell_value(iRow,28).strip()
                sServiceLevel=sheet1.cell_value(iRow,30).strip()
                sTicketType=sheet1.cell_value(iRow,29).strip()

                sSICode=sServiceType+"_"+sServiceSubType+"_"+sServiceLevel+"_"+sTicketType
                #print('SICode：'+sSICode)

                
                if sSICode in cols:
                    i=cols.index(sSICode)
                    if i>=0:
                        tempRows=sheet3.row_values(i)
                        #print("行："+str(i))
                        #print(tempRows)
                        sServiceItemNo=tempRows[0]
                        priceItem=tempRows[iPriceCol]
                        ws.write(iRow2,46,sSICode)
                        ws.write(iRow2,47,sServiceItemNo)
                        ws.write(iRow2,48,priceItem)
                else:
                    print('Not in Cols')
            else:
                print('Row'+str(iRow)+'TicketNo:'+sTicketNo +'Area:'+sArea2)
              
        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="7.ComputePrice1_"+strTime+".xls"
        wb.save(sPath+sReportName)        
       

        print("Work is over.")
        tkinter.messagebox.showinfo('提示','处理完毕，生成文件：\n'+sPath+sReportName)   
        

    
    #初始化
    root=Tk()

    #设置窗体标题
    root.title('Compute Price')

    #设置窗口大小和位置
    root.geometry('660x300+570+200')


    lbl_Summary=Label(root,text='Summary File:',width=16,justify = tkinter.RIGHT,)
    txt_SummaryFile=Entry(root,bg='white',width=68)
    btn_browse1=Button(root,text='Browse',width=8,command=selectSummaryFile)

    lbl_Prices=Label(root,text='Price File:',width=16,justify = tkinter.RIGHT)
    txt_PriceFile=Entry(root,bg='white',width=68)
    btn_browse2=Button(root,text='Browse',width=8,command=selectPriceFile)

    lbl_Area=Label(root,text="Select Area:",width=16,justify = tkinter.RIGHT)

    sArea = tkinter.StringVar()
    lst_Area=ttk.Combobox(root,width=24,textvariable=sArea) #下拉列表框
    lst_Area['values']=('AU','NZ','CN','HK','TW','SG','TH','ID','MY','VN','PH')
 

    lbl_Quarter=Label(root,text="Select Quarter:",justify = tkinter.RIGHT)

    r = tkinter.StringVar()
    r.set("Q1")
    
    rdo_Quarter1 = tkinter.Radiobutton(root,
                                      variable = r,
                                      value = "Q1",
                                      text = "Q1")
    rdo_Quarter2 = tkinter.Radiobutton(root,
                                      variable = r,
                                      value = "Q2",
                                      text = "Q2")
    rdo_Quarter3 = tkinter.Radiobutton(root,
                                      variable = r,
                                      value = "Q3",
                                      text = "Q3")
    rdo_Quarter4 = tkinter.Radiobutton(root,
                                      variable = r,
                                      value = "Q4",
                                      text = "Q4")

    
    btn_process=Button(root,text='Pocess',width=8,command=doProcess)
    btn_exit=Button(root,text='Exit',width=8,command=closeThisWindow)
 

    lbl_Summary.pack()
    txt_SummaryFile.pack()
    btn_browse1.pack()

    lbl_Prices.pack()
    txt_PriceFile.pack()
    btn_browse2.pack()



    lbl_Area.pack()

    lbl_Quarter.pack()
    rdo_Quarter1.pack()
    rdo_Quarter2.pack()
    rdo_Quarter3.pack()
    rdo_Quarter4.pack()
    
    
    btn_process.pack()
    btn_exit.pack() 

    lbl_Summary.place(x=30,y=30)
    txt_SummaryFile.place(x=130,y=30)
    btn_browse1.place(x=550,y=26)

    lbl_Prices.place(x=40,y=60)
    txt_PriceFile.place(x=130,y=60)
    btn_browse2.place(x=550,y=56)

    lbl_Area.place(x=36,y=90)
    lst_Area.place(x=130,y=90)

    lbl_Quarter.place(x=310,y=90)
    rdo_Quarter1.place(x=390,y=90)
    rdo_Quarter2.place(x=430,y=90)
    rdo_Quarter3.place(x=470,y=90)
    rdo_Quarter4.place(x=510,y=90)
    
    btn_process.place(x=250,y=160)
    btn_exit.place(x=350,y=160)
 
 
    root.mainloop() 

 
if __name__=="__main__":
    main()
