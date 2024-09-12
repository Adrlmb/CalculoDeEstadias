from tkinter import *

import pdfplumber
from openpyxl.reader.excel import load_workbook


class Application:
    def __init__(self, master=None):
        self.fontePadrao = ("Arial", "10")

        self.container1 = Frame(master)
        self.container1["pady"] = 10
        self.container1.pack()

        self.container2 = Frame(master)
        self.container2["padx"] = 20
        self.container2["pady"] = 5
        self.container2.pack()

        self.container3 = Frame(master)
        self.container3["padx"] = 20
        self.container3["pady"] = 5
        self.container3.pack()

        self.container4 = Frame(master)
        self.container4["padx"] = 20
        self.container4["pady"] = 5
        self.container4.pack()

        self.container5 = Frame(master)
        self.container5["padx"] = 20
        self.container5["pady"] = 5
        self.container5.pack()

        self.container6 = Frame(master)
        self.container6["padx"] = 20
        self.container6["pady"] = 5
        self.container6.pack()

        self.container7 = Frame(master)
        self.container7["padx"] = 20
        self.container7["pady"] = 5
        self.container7.pack()

        self.container8 = Frame(master)
        self.container8["padx"] = 20
        self.container8["pady"] = 5
        self.container8.pack()

        self.container9 = Frame(master)
        self.container9["padx"] = 20
        self.container9["pady"] = 5
        self.container9.pack()

        self.container10 = Frame(master)
        self.container10["padx"] = 20
        self.container10["pady"] = 15
        self.container10.pack()

        wb = load_workbook('estadia\Cálculo estadia.xlsx')  # Carrega o arquivo existente
        planilha = wb.active  # Seleciona a planilha ativa

        pdf = pdfplumber.open('estadia\TICKET 1.pdf')
        page = pdf.pages[0]
        text = page.extract_text()

        nomeTransportadora = StringVar()
        numeroNF = IntVar()
        nomeProduto = StringVar()
        pesoNF = StringVar()
        dataHoraSaida = StringVar()
        dataHoraChegada = StringVar()
        nomeFornecedor = StringVar()
        numeroCte = IntVar()
        nomeMotorista = StringVar()
        motivoEstadia = StringVar()


        def dataDeSaida():
            data = text.split('\n')[4].split(' ')[2].replace('.', '/')
            hora = text.split('\n')[4].split(' ')[3].split(':')[0]
            minutos = text.split('\n')[4].split(' ')[3].split(':')[1]

            return data + ' ' + hora + ':' + minutos

        def msg():
            nomeTransportadora.set(text.split('\n')[10].split('-')[1])
            numeroNF.set(int(text.split('\n')[7].split(':')[1]))
            nomeProduto.set(text.split('\n')[9].split('-')[1])
            pesoNF.set(text.split('\n')[5].split(": ")[1])
            dataHoraSaida.set(dataDeSaida())

        def preencherPlanilha():
            
            planilha['F3'] = str(self.inputTransportadora.get())  # Transportador
            planilha['C4'] = str(self.inputNF.get())  # NF
            planilha['C5'] = str(self.inputProduto.get())  # Produto, impor condição se for Rocha!!
            planilha['F12'] = str(self.inputPeso.get())  # Peso NF com vírgula
            planilha['B15'] = str(self.inputDataHoraSaida.get())  # Data e Hora da saída no formato DD/MM/AA HH:MM
            planilha['B9'] = str(self.inputDataHoraChegada.get())  # Data e Hora de chegada no formato DD/MM/AA HH:MM
            planilha['C3'] = str(self.inputFornecedor.get())  # Fornecedor
            planilha['F4'] = str(self.inputCte.get())  # Ct-e
            planilha['F5'] = str(self.inputMotorista.get())  # Motorista
            planilha['E16'] = str(self.inputMotivo.get())  # Motivo, vai se iniciar com "motivo:" e concatenar com o real motivo da estadia

            nome = nomeMotorista

            if nome in ' ':
                self.inputMotorista['bg'] = 'pink'
                self.inputMotorista['text'] = 'Preencha todos os campos'
            else:
                self.inputMotorista['bg'] = 'white'

            valor_celula = planilha['C3'].value  # Lê o valor da célula C3
            print(valor_celula)  # Imprime o valor na tela

            wb.save('EstadiasCalculadas\Estadia - ' + nome + '.xlsx')  # Salva a planilha com o nome Estadia + o nome do motorista

        # Titulo
        self.title = Label(self.container1, text="Calculo de Estadia")
        self.title["font"] = ("Calibri", "20", "bold")
        self.title.pack()


        # Transportadora
        self.transportadora = Label(self.container2, text="Transportadora ", font=self.fontePadrao)
        self.transportadora.pack(side=LEFT)
        
        self.inputTransportadora = Entry(self.container2, textvariable=nomeTransportadora)
        self.inputTransportadora["width"] = 30
        self.inputTransportadora["font"] = self.fontePadrao
        self.inputTransportadora.pack(side=LEFT)


        # Número da NF
        self.nf = Label(self.container3, text="Número da NF ", font=self.fontePadrao)
        self.nf.pack(side=LEFT)
        
        self.inputNF = Entry(self.container3, textvariable=numeroNF)
        self.inputNF["width"] = 30
        self.inputNF["font"] = self.fontePadrao
        self.inputNF.pack(side=LEFT)

        # Nome do Produto
        self.produto = Label(self.container4, text="Produto ", font=self.fontePadrao)
        self.produto.pack(side=LEFT)

        self.inputProduto = Entry(self.container4, textvariable=nomeProduto)
        self.inputProduto["width"] = 30
        self.inputProduto["font"] = self.fontePadrao
        self.inputProduto.pack(side=LEFT)

        # Peso da NF
        self.pesoNF = Label(self.container5, text="Peso da NF ", font=self.fontePadrao)
        self.pesoNF.pack(side=LEFT)

        self.inputPeso = Entry(self.container5, textvariable=pesoNF)
        self.inputPeso["width"] = 30
        self.inputPeso["font"] = self.fontePadrao
        self.inputPeso.pack(side=LEFT)

        # Número do CT-e
        self.numeroCTe = Label(self.container5, text="Número do CT-e ", font=self.fontePadrao)
        self.numeroCTe.pack(side=LEFT)

        self.inputCte = Entry(self.container5)
        self.inputCte["width"] = 30
        self.inputCte["font"] = self.fontePadrao
        self.inputCte.pack(side=LEFT)

        # Data e Hora de Chegada
        self.dataHoraChegada = Label(self.container6, text="Data/Hora de Chegada (DD/MM/AAAA HH:MM) ", font=self.fontePadrao)
        self.dataHoraChegada.pack(side=LEFT)

        self.inputDataHoraChegada = Entry(self.container6, textvariable= dataHoraChegada)
        self.inputDataHoraChegada["width"] = 30
        self.inputDataHoraChegada["font"] = self.fontePadrao
        self.inputDataHoraChegada.pack(side=LEFT)

        # Data e Hora de Saída
        self.dataHoraSaida = Label(self.container6, text="Data/Hora de Saída ", font=self.fontePadrao)
        self.dataHoraSaida.pack(side=LEFT)

        self.inputDataHoraSaida = Entry(self.container6, textvariable=dataHoraSaida)
        self.inputDataHoraSaida["width"] = 30
        self.inputDataHoraSaida["font"] = self.fontePadrao
        self.inputDataHoraSaida.pack(side=LEFT)

        # Nome do Fornecedor
        self.nomeFornecedor = Label(self.container7, text="Fornecedor ", font=self.fontePadrao)
        self.nomeFornecedor.pack(side=LEFT)

        self.inputFornecedor = Entry(self.container7)
        self.inputFornecedor["width"] = 30
        self.inputFornecedor["font"] = self.fontePadrao
        self.inputFornecedor.pack(side=LEFT)

        # Nome do Motorista
        self.nomeMotorista = Label(self.container8, text="Nome do Motorista ", font=self.fontePadrao)
        self.nomeMotorista.pack(side=LEFT)

        self.inputMotorista = Entry(self.container8)
        self.inputMotorista["width"] = 30
        self.inputMotorista["font"] = self.fontePadrao
        self.inputMotorista.pack(side=LEFT)

        # Motivo da Estadia
        self.motivoEstadia = Label(self.container9, text="Motivo da Estadia ", font=self.fontePadrao)
        self.motivoEstadia.pack(side=LEFT)

        self.inputMotivo = Entry(self.container9)
        self.inputMotivo["width"] = 30
        self.inputMotivo["font"] = self.fontePadrao
        self.inputMotivo.pack(side=LEFT)

        # Button - Chama a função que extrai os campos do PDF
        self.btnBuscar = Button(self.container10, text="Extrair dados do PDF", font=self.fontePadrao, width=20, command=msg)
        self.btnBuscar.pack(side=LEFT)

        # Button - Chama a função que salva dados do input
        self.btnInput = Button(self.container10, text="Emitir Estadia", font=self.fontePadrao, width=20, command=preencherPlanilha)
        self.btnInput.pack(side=RIGHT)

root = Tk()
Application(root)
root.mainloop()

