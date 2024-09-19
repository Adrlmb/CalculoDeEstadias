from tkinter import *

janela = Tk()
janela.title("Frame")
janela.geometry()


label1 = Label(janela, text='Fornecedor')
label1.grid(row = 0 , column = 0)

input1 = Entry(janela)
input1.grid(row = 0, column = 1)

blankSpace = Label(janela, width= 10)
blankSpace.grid(row = 0, column = 2)

label2 = Label(janela, text='Produto')
label2.grid(row = 0 , column = 3)

test = 'test'
input2 = Entry(janela, textvariable= test )
input2.grid(row = 0, column = 4)

label3 = Label(janela, text= 'Data Chegada DD/MM/AAA')
label3.grid(row = 1, column = 0)

input3 = Entry(janela)
input3.grid(row = 1, column = 1)







janela.mainloop()
