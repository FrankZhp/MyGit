from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
import xlrd
import xlwt
from xlutils.copy import copy
import os
import time
import datetime


#模块功能：从Invoice文件里读取Invoice的信息到已经汇总Mapping过的Summary文件里
def main():
    def selectExcelfile1():
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '*.xls'), ('All Files', '*')])
        txt_SummaryFile.insert(INSERT,sfname)

    def selectExcelfile2():
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel', '*.xls'), ('All Files', '*')])
        txt_InvoiceFile.insert(INSERT,sfname)

    def closeThisWindow():
        root.destroy()

    #第一个文件是Mapping Excel
    #第二个文件是Summary_Mod.xls
    def doProcess():
        tkinter.messagebox.showinfo('提示','处理Excel文件的示例程序。')
        sMergedFile=text1.get()
        sClosedFile=text2.get()
        sModFile=os.getcwd()+"\\Mod\\Summary_Mod.xls"

        print(sModFile)
        wb1 = xlrd.open_workbook(filename=sMergedFile)
        wb2 = xlrd.open_workbook(filename=sClosedFile, formatting_info=True)
        wb3 = xlrd.open_workbook(filename=sModFile)                                 #Summary Mod Excel

        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows

        sheet2 = wb2.sheet_by_index(0)   
        nrows2 = sheet2.nrows
        ncols2 = sheet2.ncols

        wb4 = copy(wb2)
        ws4 = wb4.get_sheet(0)

        #填写生成文件从AA列开始的台头
        #for i in range(ncols2):
        #     ws3.write(0,i+26,sheet2.cell_value(0,i))

        for i in range(1,nrows1):
            sTicketNo1=sheet1.cell_value(i,0)
            ws4.write(i,0,sheet1.cell_value(i,0))
            #TicketNo
            #B列Period,格式：yyyyQX,来自closedDate
            #sClosedDate=sheet1.cell_value(i,32)
            #print(sClosedDate,getPeriod(sClosedDate))
   
            #if is_valid_date(sClosedDate):
            #    ws4.write(i,1,getPeriod(sClosedDate))
            #    ws4.write(i,7,getPriceTable(sClosedDate))

            ws4.write(i,8,sheet1.cell_value(i,26))                       #Priority
            ws4.write(i,9,sheet1.cell_value(i,27))                       #Ticket No
            ws4.write(i,10,sheet1.cell_value(i,28))                      #  
            ws4.write(i,11,sheet1.cell_value(i,29))                        
            ws4.write(i,12,sheet1.cell_value(i,30))
            ws4.write(i,13,sheet1.cell_value(i,31))
            ws4.write(i,14,sheet1.cell_value(i,32))
            ws4.write(i,15,sheet1.cell_value(i,33))                      #closedDate
            ws4.write(i,16,sheet1.cell_value(i,34))
            ws4.write(i,17,sheet1.cell_value(i,35))
            ws4.write(i,18,sheet1.cell_value(i,36))
            ws4.write(i,19,sheet1.cell_value(i,37))                        
            ws4.write(i,20,sheet1.cell_value(i,38))
            ws4.write(i,21,sheet1.cell_value(i,39))
            ws4.write(i,22,sheet1.cell_value(i,40))
            ws4.write(i,23,sheet1.cell_value(i,41))                      #Country

            ws4.write(i,24,sheet1.cell_value(i,1))                       #Service Type						
            ws4.write(i,25,sheet1.cell_value(i,2))                       #Service Sub-Type
            ws4.write(i,26,sheet1.cell_value(i,3))                       #Ticket Type
            ws4.write(i,27,sheet1.cell_value(i,4))                       #Service Level
            ws4.write(i,28,sheet1.cell_value(i,6))                       #Work Order No
            ws4.write(i,29,sheet1.cell_value(i,13))                      #Support Engineer Name
            ws4.write(i,30,sheet1.cell_value(i,14))                      #Resolution
        

        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="Summary_"+strTime+".xls"
        wb4.save(sPath+sReportName)        
       

        print("Work is over.")
        tkinter.messagebox.showinfo('提示','处理完毕。')

    def getPriceTable(sDate):
        ss=datetime.datetime.strptime(sDate,"%Y-%m-%d")
        return ss
         
                       
    def getPeriod(sDate):
        try:
            ss=datetime.datetime.strptime(sDate,"%Y-%m-%d")
            sMonth=ss.month
            dict={1:'Q1',2:'Q1',3:'Q1',4:'Q2',5:'Q2',6:'Q2',7:'Q3',8:'Q3',9:'Q3',10:'Q4',11:'Q4',12:'Q4'}
            sPeriod=sDate.year+dict.get(sMonth)
            return sPeriod
        except:
            print(sDate)
            return ""

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
    root.title('Mapping')

    #设置窗口大小和位置
    root.geometry('660x300+570+200')


    label1=Label(root,text='Select Summary File:')
    txt_SummaryFile=Entry(root,bg='white',width=65)
    btn_Browse1=Button(root,text='Browse',width=8,command=selectExcelfile1)

    label2=Label(root,text='Select Invoice File:')
    txt_InvoiceFile=Entry(root,bg='white',width=65)
    btn_Browse2=Button(root,text='Browse',width=8,command=selectExcelfile2)

    lbl_SheetName = Label(root,text='Sheet Name:')
    txt_SheetName = Entry(root,bg='white',width=20)
    
    btn_Process=Button(root,text='Pocess',width=8,command=doProcess)
    btn_Exit=Button(root,text='Exit',width=8,command=closeThisWindow)
 

    label1.pack()
    txt_SummaryFile.pack()
    btn_Browse1.pack()

    label2.pack()
    txt_InvoiceFile.pack()
    btn_Browse2.pack()
    
    btn_Process.pack()
    btn_Exit.pack() 

    label1.place(x=30,y=30)
    txt_SummaryFile.place(x=150,y=30)
    btn_Browse1.place(x=550,y=26)

    label2.place(x=30,y=60)
    txt_InvoiceFile.place(x=150,y=60)
    btn_Browse2.place(x=550,y=56)

    lbl_SheetName.place(x=30,y=90)
    txt_SheetName.place(x=150,y=90)
    
    btn_Process.place(x=260,y=120)
    btn_Exit.place(x=360,y=120)
 
 
    root.mainloop() 

 
if __name__=="__main__":
    main()
