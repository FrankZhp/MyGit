from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
import xlrd
import xlwt
from xlutils.copy import copy
import os
import time
import datetime

#功能模块：把合并好的发票信息追加到整理好的Summary文件里

def main():
    def selectExcelfile1():
        root.wm_attributes('-topmost',0)
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '*.xls'), ('All Files', '*')])
        txt_InvoiceFile.insert(INSERT,sfname)
        root.wm_attributes('-topmost',1)

    def selectExcelfile2():
        root.wm_attributes('-topmost',0)
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '*.xls'), ('All Files', '*')])
        txt_SummaryFile.insert(INSERT,sfname)
        root.wm_attributes('-topmost',1)

    def closeThisWindow():
        root.destroy()

    #第一个文件是合并好的Invoice Excel
    #第二个文件是Summary_Mod.xls
    def doProcess():
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示','开始处理......')
        root.wm_attributes('-topmost',1)
        sInvoiceFile=txt_InvoiceFile.get()
        sSummaryFile=txt_SummaryFile.get()
      
        wb1 = xlrd.open_workbook(filename=sInvoiceFile)
        wb2 = xlrd.open_workbook(filename=sSummaryFile, formatting_info=True)
   
        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows

        sheet2 = wb2.sheet_by_index(0)   
        nrows2 = sheet2.nrows
        ncols2 = sheet2.ncols
        cols2 = sheet2.col_values(0)    #获取第1列TicketNo内容,放入list

        wb3 = copy(wb2)
        ws3 = wb3.get_sheet(0)

        for i in range(1,nrows1):
            
            sTicketNo=sheet1.cell_value(i,0)
            
            if sTicketNo in cols2:                                          #判断TicketNo 是否在List中  
                iRow=cols2.index(sTicketNo)
                ws3.write(iRow,1,sheet1.cell_value(i,1))                    #B列：Period                   
                ws3.write(iRow,2,sheet1.cell_value(i,2))                    #C列：Invoice 
                ws3.write(iRow,3,sheet1.cell_value(i,3))                    #D列：FESTO PO
                ws3.write(iRow,4,sheet1.cell_value(i,4))                    #E列：Invoice Amount
                ws3.write(iRow,5,sheet1.cell_value(i,5))                    #F列：INV_Currency
                ws3.write(iRow,6,sheet1.cell_value(i,6))                    #G列：Remarks
           
                
            
        

        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="6.Summary_"+strTime+".xls"
        wb3.save(sPath+sReportName)        
       

        print("Work is over.")
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示','处理完毕,生成文件：\n'+sPath+sReportName)
        root.wm_attributes('-topmost',1)


    #初始化
    root=Tk()

    #设置窗体标题
    root.title('Append Invoice To Summary ----Step 6')

    #设置窗口大小和位置
    root.geometry('660x300+430+220')


    label1=Label(root,text='Invoice Excel File:')
    txt_InvoiceFile=Entry(root,bg='white',width=66)
    btn_Browse1=Button(root,text='Browse',width=8,command=selectExcelfile1)

    label2=Label(root,text='Summay Excel File:')
    txt_SummaryFile=Entry(root,bg='white',width=66)
    btn_Browse2=Button(root,text='Browse',width=8,command=selectExcelfile2)
    
    btn_Process=Button(root,text='Process',width=8,command=doProcess)
    btn_Exit=Button(root,text='Exit',width=8,command=closeThisWindow)
 

    label1.pack()
    txt_InvoiceFile.pack()
    btn_Browse1.pack()

    label2.pack()
    txt_SummaryFile.pack()
    btn_Browse2.pack()
    
    btn_Process.pack()
    btn_Exit.pack() 

    label1.place(x=32,y=30)
    txt_InvoiceFile.place(x=142,y=30)
    btn_Browse1.place(x=550,y=26)

    label2.place(x=30,y=60)
    txt_SummaryFile.place(x=142,y=60)
    btn_Browse2.place(x=550,y=56)
    
    btn_Process.place(x=260,y=100)
    btn_Exit.place(x=360,y=100)
 
 
    root.mainloop() 

 
if __name__=="__main__":
    main()
