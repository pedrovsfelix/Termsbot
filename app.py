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
        
        self.label_filelist = CTkLabel(self.root, text='', font=("Arial", 12))
        self.label_filelist.place(x=10, y=40)
        CTkButton(self.root, width=320, height=40, command=self.browserList, text='Buscar Lista', font=("Arial", 21)).place(x=20, y=70)
        
        self.label_filedocx = CTkLabel(self.root, text='', font=("Arial", 12))
        self.label_filedocx.place(x=20, y=120)
        CTkButton(self.root, width=320, height=40, command=self.browserTerms, text='Buscar Termo', font=("Arial", 21)).place(x=20, y=150)
        
        CTkButton(self.root, width=320, height=40, command=self.initBot, text='Gerar Termos', font=("Arial", 21)).place(x=20, y=240)
    
    def browserList(self):
        file = filedialog.askopenfile(initialdir="C:/Users/pedro.felix/Documents/Projects/Python/TermsBot/", title="Selecione o arquivo da lista", filetypes=(("Excel Files", "*.xlsx*"), ("All Files", "*.*")))
        if file:
            self.label_filelist.configure(text=file.name)

    def browserTerms(self):
        file = filedialog.askopenfile(initialdir="C:/Users/pedro.felix/Documents/Projects/Python/TermsBot/", title="Selecione o arquivo do termo", filetypes=(("Docx Files", "*.docx*"), ("All Files", "*.*")))
        if file:
            self.label_filedocx.configure(text=file.name)
    
    def archives(self):
        path_list = self.label_filelist.cget("text")
        path_doc = self.label_filedocx.cget("text")
        
        try:
            table = pd.read_excel(path_list)
            columns = ["NOME", "RG", "CPF", "CARGO", "EMPRESA"]
            if not all(col in table.columns for col in columns):
                messagebox.showerror("Error", "O arquivo Excel não contém as colunas necessárias.")
                return

            os.makedirs("termos", exist_ok=True)
            
            for _, line in table.iterrows():
                document = Document(path_doc)
                
                references = {
                    "NOME": line["NOME"],
                    "CARGO": line["CARGO"],
                    "EMPRESA": line["EMPRESA"],
                    "RG": f"RG {line["RG"]}",
                    "CPF": f"CPF {line["CPF"]}"
                }
                
                for paragraph in document.paragraphs:
                    for code, value in references.items():
                        if code in paragraph.text:
                            paragraph.text = paragraph.text.replace(code, value)
                
                document.save(f"Termos/Termo - {line['NOME']}.docx")
            messagebox.showinfo("Sucesso", "Termos gerados com sucesso!")
        
        except Exception as e:
            messagebox.showerror("Error", f"Ocorreu um erro: {e}")

    def initBot(self):
        path_list = self.label_filelist.cget("text")
        path_doc = self.label_filedocx.cget("text")
        
        if path_doc and path_list:
            self.archives()
        else:
            messagebox.showerror("Error", "Vocë precisa informar os arquivos.")        

if __name__ == "__main__":
    app = App()