#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Mar  5 14:03:21 2023

@author: JaySabnis
"""

import tkinter
import customtkinter as ctk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import modifed_excel_def as med


ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")



global excelFile
excelFile=None

class App(ctk.CTk):
    
    global excelFile
    
    def __init__(self):
        super().__init__()
        self.geometry("500x500")
        self.title("RTO - PostCard Generation")
        
        frame = ctk.CTkFrame(master=self, width=460, height=400)
        frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        
        
                
        def openFile():
            try:
                
             excelFile = askopenfilename(initialdir="/", title="Select a file", filetypes=[("Excel Workbook", "*.xlsx")])
            
             if excelFile:
                    global fileFlag
                    fileFlag = 1
    
                    global text_var
                    text_var = tkinter.StringVar(value=excelFile)
                    label2 = ctk.CTkLabel(master=frame,
                                         text="Selected File",
                                         width=120,
                                         height=25,
                                         text_color="white",
                                         corner_radius=8)
                    label2.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
    
                    label3 = ctk.CTkLabel(master=frame,
                                         textvariable=text_var,
                                         width=120,
                                         height=25,
                                         fg_color=("white", "gray75"),
                                         text_color="black",
                                         corner_radius=8)
                    label3.place(relx=0.5, rely=0.6, anchor=tkinter.CENTER)
                    
                    med.excel_function(excelFile)
                    #print("Updated file created sucessfully")
            except:
                messagebox.showerror("Error", "Encountred unexpected Error while opeing file.")
        
        
        def RegistrationInput():
            dialog = ctk.CTkInputDialog(text="Enter The Registration numbers that to be printed.", title="Print Pannel")
            val=dialog.get_input().upper()
            print("Number:", val)
            new_list=list(val.split(","))
            med.print_card(new_list)
            #print("List generated successfully")
            answer=messagebox.askyesno("Software System","Do you Want to Continue ??")
            #print(answer)
            if answer: 
                pass
            else:
                app.after(1500, app.destroy())
            pass
        
    
        self.button = ctk.CTkButton(master=self,
                                   width=120,
                                   height=32,
                                   border_width=0,
                                   corner_radius=8,
                                   text="Next",
                                   command=RegistrationInput)

        self.button.grid(row=1, column=0, rowspan=3, padx=20, sticky="e")


        fileButton = ctk.CTkButton(master=frame,
                                  width=120,
                                  height=32,
                                  border_width=0,
                                  corner_radius=8,
                                  text="Click to Load File",
                                  fg_color="#A934BD",
                                  hover_color="#8C319C",
                                  command=openFile)

        fileButton.place(relx=0.5, rely=0.4, anchor=tkinter.CENTER)
        





if __name__=="__main__":
    app = App()
    app.mainloop()


