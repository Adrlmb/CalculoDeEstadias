import tkinter as tk
from tkinter import filedialog


class Tela:
    def __init__(self, master):
        self.nossaTela = master

        self.barra_menu = tk.Menu(self.nossaTela)
        self.nossaTela.config(menu=self.barra_menu)

        self.barra_menu.add_command(label="Escolher PDF", command=self.abreDir)

    def abreDir(self):
        self.arquivo = filedialog.askopenfile(initialdir = "/Desktop", title = "Selecione um arquivo", filetypes=(("Arquivos PDF", "*.pdf"),("Arquivos de texto", "*.txt")))
        print(str(self.arquivo).split('\u0027')[1])


janelaRaiz = tk.Tk()
Tela(janelaRaiz)
janelaRaiz.mainloop()
