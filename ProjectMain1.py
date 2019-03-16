import tkinter
import tkinter.messagebox
import SubForm1_ReadCSV
import SubForm2_MergeFESTO
import SubForm7_ComputePrice
import SubForm3_Mapping
import SubForm8_MakeReport
import SubForm4_Summary
import SubForm5_MergeInvoice
import SubForm6_AppendInvoiceToSummary

def main():
    #设置窗体标题
    root.title('Compute Closed Tickets Price')

    #设置窗口大小和位置
    root.geometry('800x500+360+100')

    menu = tkinter.Menu(root)
    root.config(menu=menu)
    submenu1 = tkinter.Menu(menu,tearoff=0)
    submenu1.add_command(label="1.Read CSV",command=openReadCSVForm)
    submenu1.add_command(label="2.Merge FESTO Files",command=openMergeFESTOForm)
    submenu1.add_command(label="3.Mapping",command=openMappingForm)
    submenu1.add_command(label="4.Summary",command=openSummaryForm)
    submenu1.add_command(label="5.Merge Invoice",command=openMergeInvoiceForm)
    submenu1.add_command(label="6.Append Invoice to Summary",command=openAppendInvoiceToSummaryForm)
    submenu1.add_command(label="7.Compute Price" ,command=openComputePriceForm)
    submenu1.add_command(label="8.Make Report",command=openMakeReportForm)
    submenu1.add_separator()
    submenu1.add_command(label="9.Exit",command=closeThisWindow)
    menu.add_cascade(label="Work",menu=submenu1)

    submenu2 = tkinter.Menu(menu,tearoff=0)
    submenu2.add_command(label="Help",command=openHelp)
    submenu2.add_command(label="about",command=openAbout)
    menu.add_cascade(label="Help",menu=submenu2)

    #root.config(menu=menu)
    root.mainloop

 

def openReadCSVForm():
    print("Open Read CSV Form")
    SubForm1_ReadCSV.main()

def openMergeFESTOForm():
    print("openMergeFESTOForm")
    SubForm2_MergeFESTO.main()

def openMappingForm():
    print("openMappingForm")
    SubForm3_Mapping.main()
    
def openSummaryForm():
    print("openMappingForm")
    SubForm4_Summary.main()
    
def openMergeInvoiceForm():
    print("openMergeInvoiceForm")
    SubForm5_MergeInvoice.main()

def openAppendInvoiceToSummaryForm():
    print("openAppendInvoiceToSummaryForm")
    SubForm6_AppendInvoiceToSummary.main()

def openComputePriceForm():
    print("openComputePriceForm")
    SubForm7_ComputePrice.main()

def openMakeReportForm():
    print("openMakeReportForm")
    SubForm8_MakeReport.main()

def closeThisWindow():
    if  tkinter.messagebox.askokcancel('提示', '确定要退出吗？')==True:
        root.destroy()
  


def openHelp():
     tkinter.messagebox.showinfo("欢迎", "欢迎使用计算Tickets Price Tool！\nBy Frank Zhang")
  

def openAbout():
    print("openAbout")


    
if __name__=="__main__":
    root = tkinter.Tk()
    main()
                    
                    
                    
