from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
import os
import csv
from xlwt import *
import time

def main():
    def selectExcelfile():
        sfname = filedialog.askopenfilename(title='Please Select CSV File', filetypes=[('CSV', '*.csv'), ('All Files', '*')])
        #print(sfname)
        txt_CsvFile.insert(INSERT,sfname)
        root.wm_attributes('-topmost',1)


    def closeThisWindow():
        root.destroy()

    def doProcess():
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示','开始处理CSV文件......')
        root.wm_attributes('-topmost',1)

        workbook=Workbook(encoding = 'utf-8')
        worksheet = workbook.add_sheet('sheet1')

        sCsvFile=txt_CsvFile.get()
        sCurrentPath=os.getcwd()
   
        i=0
        with open(sCsvFile,newline='',encoding='UTF-8') as csvfile:
            rows=csv.reader(csvfile)
            for iRow,row in enumerate(rows):
                #判断有几列
                if iRow==1:
                    iCols=len(row)
            
                for iCol in range(0,len(row)):
                    worksheet.write(iRow,iCol,','.join(row).split(',')[iCol])

        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="1.ClosedTickets_"+strTime+".xls"            
        workbook.save(sPath+sReportName)
        print("Process is over.")
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示信息','已经转换生成Excel文件。\n '+sPath+sReportName)
        root.wm_attributes('-topmost',1)
    
    #初始化
    root=Tk()

    #设置窗体标题
    root.title('Read csv file---Step 1')

    #设置窗口大小和位置
    root.geometry('660x300+430+220')


    label1=Label(root,text='Select CSV File:')
    txt_CsvFile=Entry(root,bg='white',width=70)
    btn_Browse=Button(root,text='Browse',width=8,command=selectExcelfile)


    btn_Do=Button(root,text='Process',width=8,command=doProcess)
    btn_Exit=Button(root,text='Exit',width=8,command=closeThisWindow)

    label1.pack()
    txt_CsvFile.pack()
    btn_Browse.pack()



    btn_Do.pack()
    btn_Exit.pack()
    

    label1.place(x=30,y=30)
    txt_CsvFile.place(x=120,y=30)
    btn_Browse.place(x=550,y=26)

    
    btn_Do.place(x=230,y=120)
    btn_Exit.place(x=330,y=120)
 
    root.mainloop() 

 
if __name__=="__main__":
    main()
