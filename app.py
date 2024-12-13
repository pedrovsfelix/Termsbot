from customtkinter import *
from tkinter import messagebox, filedialog
import pandas as pd
from docx import Document
import os

class App:
    def __init__(self):
        self.root = CTk()
        self.root.title("TermsBot")
        self.root.geometry("350x300")
        self.widgets()
        self.root.mainloop()
    
    def widgets(self):
        CTkLabel(self.root, text='Gere termos de forma automática', font=("Arial", 18)).place(x=50, y=0)
        
        self.label_filelist = CTkLabel(self.root, text='', font=("Arial", 12)).place(x=10, y=40)
        CTkButton(self.root, width=320, height=40, text='Buscar Lista', font=("Arial", 21)).place(x=20, y=70)
        
        self.label_filedocx = CTkLabel(self.root, text='', font=("Arial", 12)).place(x=10, y=40)
        CTkButton(self.root, width=320, height=40, text='Buscar Termo', font=("Arial", 21)).place(x=20, y=150)
        
        CTkButton(self.root, width=320, height=40, text='Gerar Termos', font=("Arial", 21)).place(x=20, y=240)
        


if __name__ == "__main__":
    app = App()