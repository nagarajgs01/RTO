import tkinter
import customtkinter as ct
from tkinter.filedialog import askopenfile
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfile
from tkinter import messagebox

from main import ngs

pt = ngs.getBasePath()

ngs.pathAloc(pt)
# ngs.createDir(ngs.output)
# ngs.readXL(ngs.excel)
# ch = input("Enter Choice :")
# if ch == "1":
#     rng=input("Enter Range : ")
#     ngs.rngSelected(rng)
# elif ch == "2":
#     ngs.allSelected()
# else:
#     print("Invalid Choice")



from time import sleep
import trace

ct.set_appearance_mode("dark")
ct.set_default_color_theme("green")

global f_var
global excelFile
global fileFlag
global text_var
global rangeEntry

global excel
global word
global output

fileFlag = 0
excelFile = None
f_var = 0

""" Second Frame to choose range or all files """


class secondFrame(ct.CTkFrame):

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.sel_op()

        # path of file
        global excelFile
        text_var = excelFile
        print(text_var)

    def sel_op(self):
        radio_var = tkinter.IntVar(0)
        ngs.createDir(ngs.output)
        ngs.readXL(ngs.excel)

        # radio button fuction
        def radiobutton_event():
            print("radiobutton toggled, current value:", radio_var.get())
            val = int(radio_var.get())
            # range code
            if val == 1:
                rangeEntry.configure(state="normal")
                rangeEntry.focus()


            # select all
            if val == 2:
                rangeEntry.select_clear()
                ngs.allSelected()
                rangeEntry.configure(state="disabled")

        self.label = ct.CTkLabel(self, text="Note : Enter Range example 1-10, 15-20 .....")
        self.label.grid(row=0, column=0, padx=20)
        radiobutton_1 = ct.CTkRadioButton(master=self, text="Selecte Range",
                                          command=radiobutton_event, variable=radio_var, value=1)
        radiobutton_2 = ct.CTkRadioButton(master=self, text=" Select All ",
                                          command=radiobutton_event, variable=radio_var, value=2)
        global rangeEntry
        rangeEntry = ct.CTkEntry(master=self,
                            placeholder_text="Enter Range 1-10,15-20...",
                            width=120,
                            height=25,
                            border_width=2,
                            corner_radius=10,
                            state="disabled")

        rangeEntry.place(relx=0.5, rely=0.4, anchor=tkinter.CENTER)

        radiobutton_1.place(relx=0.5, rely=0.3, anchor=tkinter.CENTER)
        radiobutton_2.place(relx=0.5, rely=0.6, anchor=tkinter.CENTER)


"""First Frame to load file """


class firstFrame(ct.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.label = ct.CTkLabel(self, text="Note : Select the Excel Workbook file ...")
        self.label.grid(row=0, column=0, padx=20)

        fileButton = ct.CTkButton(master=self,
                                  width=120,
                                  height=32,
                                  border_width=0,
                                  corner_radius=8,
                                  text="Click to Load File",
                                  fg_color="#A934BD",
                                  hover_color="#8C319C",
                                  command=self.open_file)

        fileButton.place(relx=0.5, rely=0.4, anchor=tkinter.CENTER)

    def open_file(self):
        try:
            # Loading excel file path variable
            global excelFile
            excelFile = None
            # fetches excel file path
            excelFile = askopenfilename(initialdir="/", title="Select a file", filetypes=[("Excel Workbook", "*.xlsx")])
            # check if file is blank or not

            ngs.xlPath(excelFile)


            if excelFile:
                global fileFlag
                fileFlag = 1

                global text_var
                text_var = tkinter.StringVar(value=excelFile)
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
            messagebox.showerror("Error", "Encountred unexpected Error while opeing file.")


"""root window , Consist of buttons and frame calling"""


class App(ct.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("600x500")
        self.title("System Report Generation")
        self.resizable(width=False, height=False)

        self.first_frame()

    # configure Button Functions based on First or Second Frame

    def configureButton(self):

        def save_file():
            global rangeEntry
            ngs.rngSelected(rangeEntry.get())
            messagebox.showinfo("success","Successfull")
            self.after(1000,self.destroy())

        print("Inside frame ", f_var)
        global fileFlag

        if f_var == 1:
            self.button.configure(command=self.second_frame)

        if f_var == 2:
            self.button.configure(text="Finish")
            self.button.configure(command=save_file)

    def first_frame(self):

        global f_var
        f_var = 1

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=9)

        # button
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self.my_frame = firstFrame(master=self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        global button

        self.button = ct.CTkButton(master=self,
                                   width=120,
                                   height=32,
                                   border_width=0,
                                   corner_radius=8,
                                   text="Next")

        self.button.grid(row=1, column=0, rowspan=3, padx=20, sticky="e")

        self.configureButton()

    def second_frame(self):

        def clean():
            # destory Back button and goes to first frame from second frame on click
            global fileFlag
            fileFlag = 0
            backButton.destroy()
            self.first_frame()

        global f_var
        f_var = 2

        self.my_frame = secondFrame(self)
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
                                  command=clean)
        backButton.grid(row=1, column=0, rowspan=3, padx=20, sticky="w")
        self.configureButton()


app = App()
app.mainloop()
