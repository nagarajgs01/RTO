# -*- coding: utf-8 -*-
"""
Created on Wed Mar 15 16:45:59 2023

@author: NAGARAJ G SHEELI
"""

import tkinter as tk
from tkinter import *
from tkinter import ttk
from PIL import Image, ImageTk
from Interface import App

import os

root=Tk()
kaLogo=PhotoImage(file="kalogo.png")
rtoLogo=PhotoImage(file="rtologo.png")
g20Logo=PhotoImage(file="glogo.png")
makeLogo=PhotoImage(file="makeinindialogo.png")
embLogo=PhotoImage(file="emblem.png")

height=430
width=530
x=(root.winfo_screenwidth()//2)-(width//2)
y=(root.winfo_screenheight()//2)-(height//2)


root.geometry('{}x{}+{}+{}'.format(width,height,x,y))
root.overrideredirect(True)


root.config(background="#ffffff")


madeby=Label(text="Made By - \nNagaraj G Sheeli  &  Prabhakar M Sabnis",bg='#ffffff',font=('Comic Sans MS',10,"bold"), fg="#000000")
madeby.place(x=140,y=380)

kalogo=Label(root,image=kaLogo)
kalogo.place(x=200,y=85)

rtologo=Label(root,image=rtoLogo)
rtologo.place(x=100,y=220)

g20logo=Label(root,image=g20Logo)
g20logo.place(x=227,y=10)

makelogo=Label(root,image=makeLogo)
makelogo.place(x=430,y=10)

emblogo=Label(root,image=embLogo)
emblogo.place(x=20,y=10)

root.after(3000,App)


#root.resizable(False,False)
root.mainloop()



