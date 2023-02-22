import tkinter
import customtkinter as ct
from tkinter.filedialog import askopenfile
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfile
from tkinter import messagebox



import pandas as pd
import openpyxl
from datetime import date
from datetime import datetime
from pathlib import Path
from docxtpl import DocxTemplate


ct.set_appearance_mode("dark")
ct.set_default_color_theme("green")

date = date.today()


global button
global f_var
global file
f_var=0




#chosing frame
class MyFrame2(ct.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.sel_op()
        
    

    def ngs(self,file):
        base_dir = Path(__file__).parent
        word = "F:/7th Sem/RTO Project/Book1.docx"
        excel = file #file
        
        #code after clicking save
        
        output = base_dir / "OUTPUT" # path of the file to be saved

        output.mkdir(exist_ok=True)
    
        fd = pd.read_excel(excel, 'Sheet1')
        
        
        for record in fd.to_dict("records"):
            doc = DocxTemplate(word)
            doc.render(record)
            output_path = output / f"{record['Name']}-doc.docx"
            doc.save(output_path)

        
        
        
    def sel_op(self):    
        radio_var = tkinter.IntVar(0)
        def radiobutton_event():
                print("radiobutton toggled, current value:", radio_var.get())
                val=int(radio_var.get())
                if val==1:
                    entry.configure(state="normal")
                    entry.focus()
                    
                if val==2:
                    entry.select_clear()
                    entry.configure(state="disabled")
                    global file
                    self.ngs(file)
                    
                
        self.label = ct.CTkLabel(self,text="Note : Enter Range example 1-10, 15-20 .....")
        self.label.grid(row=0, column=0, padx=20)
        radiobutton_1 = ct.CTkRadioButton(master=self, text="Selecte Range",
                                             command=radiobutton_event, variable= radio_var, value=1)
        radiobutton_2 = ct.CTkRadioButton(master=self, text=" Select All ",
                                             command=radiobutton_event, variable= radio_var, value=2)
        
        entry = ct.CTkEntry(master=self,
                               placeholder_text="Enter Range 1-10,15-20...",
                               width=120,
                               height=25,
                               border_width=2,
                               corner_radius=10,
                               state="disabled")
        
        entry.place(relx=0.5, rely=0.4, anchor=tkinter.CENTER)
        
        radiobutton_1.place(relx=0.5, rely=0.3,anchor=tkinter.CENTER)
        radiobutton_2.place(relx=0.5, rely=0.6,anchor=tkinter.CENTER) 
        
    
            
        
        
  
       



#initial frame
class MyFrame(ct.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.label = ct.CTkLabel(self,text="Note : Select the Excel Workbook file ...")
        self.label.grid(row=0, column=0, padx=20)
        
        fileButton= ct.CTkButton(master=self,
                                 width=120,
                                 height=32,
                                 border_width=0,
                                 corner_radius=8,
                                 text="Click to Load File",
                                 fg_color="#A934BD",
                                 hover_color="#8C319C",
                                 command=self.open_file)
        fileButton.place(relx=0.5, rely=0.4,anchor=tkinter.CENTER)
        

        
    def open_file(self):
        try:
            global file
            file=None
        #file = askopenfile(mode="r",filetypes=[("Excel Workbook","*.xlsx")])
            file = askopenfilename(initialdir="/",title="Select a file",filetypes=[("Excel Workbook","*.xlsx")])
            if file :
                text_var = tkinter.StringVar(value=file)
                label2 = ct.CTkLabel(master=self,
                                   text="Selected File",
                                   width=120,
                                   height=25,
                                   text_color="white",
                                   corner_radius=8)
                label2.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
                
                label3 = ct.CTkLabel(master=self,
                                   textvariable=text_var,
                                   width=120,
                                   height=25,
                                   fg_color=("white", "gray75"),
                                   text_color="black",
                                   corner_radius=8)
                label3.place(relx=0.5, rely=0.6, anchor=tkinter.CENTER)
        except:
            messagebox.showerror("Error","Encountred unexpected Error while opeing file.")
        
        
   
        
        
    


class App(ct.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("600x500")
        self.title("System Report Generation")
        self.resizable(width=False,height=False)
       
       
        self.back_frame()
              
    def test(self):
        
        def save_file():
            print("file save")
            files = [('All Files', '*.*'), 
                     ('Python Files', '*.py'),
                     ('Text Document', '*.txt')]
            
            file = asksaveasfile(mode='w',filetypes = files, defaultextension = files)
            
            
            
            
        print(f_var)
        if f_var==1:
            self.button.configure(command=self.next_frame)
            
        if f_var==2:
            self.button.configure(command=save_file)
           
            
    def back_frame(self):
        #frame
        global f_var
        f_var=1
        self.columnconfigure(0,weight=1)
        self.rowconfigure(0,weight=9)
        
        #button
        self.columnconfigure(0,weight=1)
        self.rowconfigure(1,weight=1)

        self.my_frame = MyFrame(master=self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        global button
        self.button = ct.CTkButton(master=self,
                                 width=120,
                                 height=32,
                                 border_width=0,
                                 corner_radius=8,
                                 text="Next",
                                 command=self.next_frame)
        self.button.grid(row=1,column=0,rowspan=3,padx=20,sticky="e")
        self.test()
   
        
        
        
        
    def next_frame(self) :
        
        def con():
            backButton.destroy()
            self.back_frame()
            
            
        global f_var
        f_var=2
        
        self.my_frame= MyFrame2(self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        backButton = ct.CTkButton(master=self,
                                 width=120,
                                 height=32,
                                 border_width=1,
                                 corner_radius=8,
                                 text="Back",
                                 fg_color="transparent",
                                 hover_color="grey",
                                 border_color="white",
                                 command=con)
        backButton.grid(row=1,column=0,rowspan=3,padx=20,sticky="w")
        self.test()
        
            

app=App()
app.mainloop()
