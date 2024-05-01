from tkinter import *
from tkinter import filedialog
#from tkinter.messagebox import showinfo
from tkinter import messagebox
import pandas as pd
import os
import numpy as np 
import openpyxl
from openpyxl import load_workbook
from tkinter.filedialog import askdirectory


TestResultCheck = 2

FileName1 = ""
FileName2 = ""
    
def browseTestReport():
    global Report_Path
    global TestResultCheck
    TestResultCheck = os.system('start "excel.exe" "TestReport.xlsx"')
    Report_Path = "TestReport.xlsx"
    

def RB_Command():      
    My_Str2.set("")
    
    if(RB_v.get() == 1):
        My_Str3.set("Select Comparison Report:")
        button_explore2.place(x=275, y=326)
        
    else:
        My_Str3.set("Select Path to Save Report:")
        button_explore2.place(x=275, y=326)
     
       
   
    #Button_Value1 = RB_v.get()


def browseComparisonReport():
    global FileName1
    global FileName2
    
    if(RB_v.get() == 1):       
        FileName2 = filedialog.askopenfilename(initialdir = "/",title = "Select Comparison Report",filetypes = (("Excel file","*.xlsx*"),("All files","*.*")))
        # Change label contents
        My_Str2.set(FileName2)
        FileName1 = ""
        
    
    if(RB_v.get() == 2):
        global Report_Folder        
        FileName1 = askdirectory(title='Select Folder') # shows dialog box and return the path
        #print(path)
        My_Str2.set(FileName1)
        Report_Folder = FileName1
        FileName2 = ""

def ExitTool():
    res1=messagebox.askquestion('Exit Application', 'Do you really want to exit')
    if res1 == 'yes' :
        window.destroy()
    else :
        messagebox.showinfo('Return', 'Returning to main tool')

        
def ResetWindow():
    res2=messagebox.askquestion('Reset Application', 'Do you really want to reset window')
    if res2 == 'yes' :
        global FileName1
        global FileName2
        #global My_Str1
        global My_Str2
        global RB_v
        
        FileName1 = ""
        FileName2 = ""      
        #My_Str1.set("")
        My_Str2.set("")
        My_Str3.set("")
        RB_v.set(0)        
        
        entry1.delete(0, END)
        entry2.delete(0, END)
        entry3.delete(0, END)
        
        button_explore2.place_forget()
        
        
                       
    else :
        messagebox.showinfo('Return', 'Returning back to tool')
    
    
def CreateReport():
    global FileName1
    global FileName2
    global TestResultCheck
    #print(test101)
    Comparison_Path = FileName2    
    #Report_Folder = os.path.dirname(Report_Path)
    ProjectName=entry1.get()
    TestType = entry2.get()
    TestDate = entry3.get()
    
    if len(ProjectName) == 0:
        messagebox.showwarning("Warning", "Please Enter Project Name")
        return -1
    if len(TestType) == 0:
        messagebox.showwarning("Warning", "Please Enter Test Name")
        return -1
    if len(TestDate) == 0:
        messagebox.showwarning("Warning", "Please Enter Test Date")
        return -1
    
    if (TestResultCheck == 1):
        messagebox.showwarning("Warning", "Browser Test Result")
        return -1
        
    #print(RadioFlag)
    
    if (RB_v.get() == 1):
        #print(RadioFlag)
        if len(FileName2) == 0:
            messagebox.showwarning("Warning", "Please Select the Comparison Report")
            return -1
        
    elif (RB_v.get() == 2):
        if len(FileName1) == 0:
            messagebox.showwarning("Warning", "Select Report Folder")
            return -1
        
    else:       
        messagebox.showwarning("Warning", "Select Yes or No for Comparison Report")
        return -1
    
    
    TestDate = "_" + TestDate
    ColumnName1 = "_" + TestType + TestDate
        
    
    if(len(Comparison_Path) == 0):
        Comparison_Path = Report_Folder + '\\' + ProjectName + 'PerfTestReport.xlsx'
        #Create new workbook
        wb = openpyxl.Workbook()
        #Get SHEET name
        ws = wb.active
        ws.title = "Comparison"
        #Sheet name = wb.sheetnames
        #Save created workbook at same path where .py file exist
        wb.save(filename=Comparison_Path)
        df3 = pd.read_excel(Report_Path, usecols=['Label'])
        df3.to_excel(Comparison_Path, sheet_name = 'Comparison', index=False)
    
#-----------DataFrame creation for comparison sheet---------------------------#
    
    # Creating dataframe for Comparison Sheet
    df1= pd.read_excel(Comparison_Path, sheet_name = 'Comparison')
    df2 = pd.read_excel(Report_Path, usecols=['Label', 'Average', '90th pct'])
    df2.columns = df2.columns.map(lambda x: x+ColumnName1 if x !='Label' else x)
    SimpleDataFramel = pd.merge(df1,df2, on="Label", how="outer")


#---------------------------Generating a comparison report----------------------------------#

    #Generating workbook
    ExcelWorkbook = load_workbook(Comparison_Path)
    #Generating the writer engine
    writer = pd.ExcelWriter(Comparison_Path, engine = 'openpyxl')
    #Assigning the workbook to the writer engine
    writer.book = ExcelWorkbook
    writer.sheets = dict((ws.title, ws) for ws in ExcelWorkbook.worksheets)
    SimpleDataFramel.to_excel(writer, sheet_name = 'Comparison', index=False, startcol=0, startrow=0)
    # Adding the DataFrames to the excel as a new sheet
    #SimpleDataFramel.to_excel(writer, sheet name 'Comparison', index False)
    TestResultCheck = 1
    writer.save()
    writer.handles = None
    


#---------------------------Adding New test result to comparison report------------------------------#


    #Generating workbook
    
    sheetname= TestType + TestDate
    ExcelWorkbook = load_workbook(Comparison_Path)
    #Generating the writer engine
    writer = pd.ExcelWriter(Comparison_Path, engine = 'openpyxl')
    #Assigning the workbook to the writer engine
    writer.book = ExcelWorkbook            
    #Creating first dataframe
    SimpleDataFrame2=pd.read_excel(Report_Path)            
    #print(SimpleDataFrame)            
    #Adding the DF to the excel as a new sheet
    SimpleDataFrame2.to_excel(writer, sheet_name = sheetname, index=False)
    writer.save()
    writer.handles =None
    messagebox.showinfo("Info", "Report Created Successfully")
    


#---------------------Window Setup--------------------------------#

# Create the root window
window = Tk()
#window = Frame(root)
Welcome_Color = "red"
BG_Colour1 = "gray78"
FG_Colour1 = "gray20"
BG_Button = "dimgray"
FG_Button = "floralwhite"
Link_Colour = 'violetred4'

# Set window title
window.title('JRCT')
# Set window size
window.minsize(width =700, height =550)
window.maxsize(width =900, height =620)
# Set window Backgroung color
window.configure(bg=BG_Colour1)

# Create a window label
label_file_explorer = Label(window,text = "Welcome to Jmeter Report Comparison Tool", font = ("Arial", 10, "bold"), fg = Welcome_Color, bg=BG_Colour1)
label_file_explorer.pack()

Label1 = Label(window,text = 'Project Name:', font = ("Arial", 9, "bold"), fg=FG_Colour1, bg=BG_Colour1)
Label1.place(x=40, y=75)

entry1 = Entry(window,bd =4)
entry1.place(x=150, y=75)

Label2 = Label(window,text = 'Test Name:', font = ("Arial", 9, "bold"), fg=FG_Colour1, bg=BG_Colour1)
Label2.place(x=40, y=105)

entry2 = Entry(window,bd =4)
entry2.place(x=150, y=105)

Label3 = Label(window,text = 'Test Date:', font = ("Arial", 9, "bold"), fg=FG_Colour1, bg=BG_Colour1)
Label3.place(x=40, y=135)

entry3 = Entry(window,bd = 4)
entry3.place(x=150, y=135)

#My_Str1 = StringVar()
#My_Str1.set("")
#filepath1 = Label(window,textvariable = My_Str1 , anchor= W, bg=BG_Colour1, fg = Link_Colour)
#filepath1.place(x=250, y=215)

My_Str2 = StringVar()
My_Str2.set("")
filepath2 = Label(window,textvariable = My_Str2 , anchor= W, bg=BG_Colour1, fg = Link_Colour)
filepath2.place(x=275, y=360)


button_explore1 = Button(window,text = "Paste Test Result Here", font = ("Arial", 10, "bold"), fg= FG_Button, bg=BG_Button, command = browseTestReport)
button_explore1.place(x=40, y=215)


button_explore2 = Button(window, text = "Browse", font = ("Arial", 10, "bold"), fg= FG_Button, bg=BG_Button, command = browseComparisonReport)

#button_explore3 = Button(window, text = "Select Folder For Report", font = ("Arial", 10, "bold"), fg= FG_Button, bg=BG_Button, command = browseComparisonReport)

LabelRB = Label(window,text = 'Already have a comparison report?', font = ("Arial", 9, "bold"), fg=FG_Colour1,bg=BG_Colour1)
LabelRB.place(x=40, y=275)

My_Str3 = StringVar()
My_Str3.set("")
filepath3 = Label(window,textvariable = My_Str3 , anchor= W, font = ("Arial", 9, "bold"), fg=FG_Colour1,bg=BG_Colour1)
filepath3.place(x=40, y=330)

global RB_v
RB_v = IntVar()

RB1 = Radiobutton(window, text='Yes', variable=RB_v, value=1, bg=BG_Colour1, command = RB_Command)
RB1.place(x=270, y=275)

RB2 = Radiobutton(window, text='No', variable=RB_v, value=2, bg=BG_Colour1, command = RB_Command)
RB2.place(x=350, y=275)

bottomframe = Frame(window, bg=BG_Colour1)
bottomframe.pack( side = BOTTOM, pady=40)

button_exit = Button(bottomframe,text = "Exit", font = ("Arial", 10, "bold"), fg= FG_Button, bg=BG_Button, command = ExitTool)
button_exit.pack( side = RIGHT, padx=55)

button_reset = Button(bottomframe,text = "Reset", font = ("Arial", 10, "bold"), fg= FG_Button, bg=BG_Button, command = ResetWindow)
button_reset.pack( side = LEFT, padx=55)

button_Generate_Report = Button(bottomframe,text = "Generate Report", font = ("Arial", 10, "bold"), fg= FG_Button, bg=BG_Button, command = CreateReport)
button_Generate_Report.pack( side = TOP)

# Let the window wait for any events
window.mainloop()
