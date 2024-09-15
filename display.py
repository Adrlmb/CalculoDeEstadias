from tkinter import *
from tkinter import filedialog
import pdfplumber
from openpyxl.reader.excel import load_workbook


class Application:
    def __init__(self, master=None):

        def dataDeSaida(text):
            data = text.split('\n')[4].split(' ')[2].replace('.', '/')
            hora = text.split('\n')[4].split(' ')[3].split(':')[0]
            minutos = text.split('\n')[4].split(' ')[3].split(':')[1]

            return data + ' ' + hora + ':' + minutos

        def escolherPdf():
            self.pdf = filedialog.askopenfile(initialdir="/Desktop", title="Selecione um arquivo", filetypes=[("Arquivos PDF", "*.pdf")])
            if self.pdf:
                return str(self.pdf).split('\u0027')[1]

        def importarPdf():
            pdf = pdfplumber.open(escolherPdf())
            page = pdf.pages[0]
            text = page.extract_text()

            nomeTransportadora.set(text.split('\n')[10].split('-')[1])
            numeroNF.set(text.split('\n')[7].split(':')[1])
            nomeProduto.set(text.split('\n')[9].split('-')[1])
            pesoNF.set(text.split('\n')[5].split(": ")[1])
            dataHoraSaida.set(dataDeSaida(text))


        def conferenciaDeDados():

            campoTransportadora = str(self.inputTransportadora.get().upper())
            campoNumeroNF = str(self.inputNF.get())  # NF
            campoProduto = str(self.inputProduto.get().upper())  # Produto, impor condição se for Rocha!!
            campoPesoNF = str(self.inputPeso.get())  # Peso NF com vírgula
            campoSaida = str(self.inputDataHoraSaida.get())  # Data e Hora da saída no formato DD/MM/AA HH:MM
            campoEntrada = str(self.inputDataHoraChegada.get())  # Data e Hora de chegada no formato DD/MM/AA HH:MM
            campoFornecedor = str(self.inputFornecedor.get().upper())  # Fornecedor
            campoCte = str(self.inputCte.get())  # Ct-e
            campoMotorista = str(self.inputMotorista.get().upper())  # Motorista
            campoMotivos = 'MOTIVO: ' + str(self.inputMotivo.get().upper())  # Motivo, vai se iniciar com "MOTIVO: " e concatenar com o real motivo da estadia


            campos = [campoFornecedor, campoTransportadora, campoMotorista, campoProduto, campoEntrada, campoSaida, campoCte, campoNumeroNF, campoPesoNF, campoMotivos]

            fields = [self.inputFornecedor, self.inputTransportadora, self.inputMotorista, self.inputProduto, self.inputDataHoraChegada, self.inputDataHoraSaida, self.inputCte, self.inputNF, self.inputPeso, self.motivoEstadia]

            for i in range(10):
                if campos[i] in '':
                    fields[i].focus()
                    fields[i]['bg'] = 'pink'
                else:
                    fields[i]['bg'] = 'white'


        def preencherPlanilha():
            wb = load_workbook('estadia\Cálculo estadia.xlsx')  # Carrega o arquivo existente
            planilha = wb.active  # Seleciona a planilha ativa

            conferenciaDeDados()

            # Pega os dados dos inputs e coloca na planilha de acordo com a célula referenciada
            planilha['F3'] = str(self.inputTransportadora.get().upper())  # Transportador
            planilha['C4'] = str(self.inputNF.get())  # NF
            planilha['C5'] = str(self.inputProduto.get().upper())  # Produto, impor condição se for Rocha!!
            planilha['F12'] = str(self.inputPeso.get())  # Peso NF com vírgula
            planilha['B15'] = str(self.inputDataHoraSaida.get())  # Data e Hora da saída no formato DD/MM/AA HH:MM
            planilha['B9'] = str(self.inputDataHoraChegada.get())  # Data e Hora de chegada no formato DD/MM/AA HH:MM
            planilha['C3'] = str(self.inputFornecedor.get().upper())  # Fornecedor
            planilha['F4'] = str(self.inputCte.get())  # Ct-e
            planilha['F5'] = str(self.inputMotorista.get().upper())  # Motorista
            planilha['E16'] = 'MOTIVO: ' + str(self.inputMotivo.get().upper())  # Motivo, vai se iniciar com "MOTIVO: " e concatenar com o real motivo da estadia

            salvarPlanilha(wb)

        def salvarPlanilha(wb):
            # if nome in ' ':
            #     self.inputMotorista['bg'] = 'pink'
            #     self.inputMotorista['text'] = 'Preencha todos os campos'
            # else:
            #     self.inputMotorista['bg'] = 'white'

            self.planilha = filedialog.asksaveasfilename(initialdir="C:/Users/adrie/PycharmProjects/pythonProject1/EstadiasCalculadas", defaultextension=".xlsx",
                                                         title="Salvar como", filetypes=[("Excel", "*.xlsx")])

            if self.planilha:
                wb.save(self.planilha)
                wb.close()

        def limparDados():
            self.inputTransportadora.delete(0, END)
            self.inputNF.delete(0, END)
            self.inputProduto.delete(0, END)
            self.inputPeso.delete(0, END)
            self.inputDataHoraSaida.delete(0, END)
            self.inputDataHoraChegada.delete(0, END)
            self.inputFornecedor.delete(0, END)
            self.inputCte.delete(0, END)
            self.inputMotorista.delete(0, END)
            self.inputMotivo.delete(0, END)

            self.inputFornecedor.focus()# Puxa o foco para o campo Fornecedor

        nomeTransportadora = StringVar()
        numeroNF = StringVar()
        nomeProduto = StringVar()
        pesoNF = StringVar()
        dataHoraSaida = StringVar()
        dataHoraChegada = StringVar()
        nomeFornecedor = StringVar()
        numeroCte = StringVar()
        nomeMotorista = StringVar()
        motivoEstadia = StringVar()

        # -----------------------------------------------------------------------------------------

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

        # Titulo
        self.title = Label(self.container1, text="Calculo de Estadia")
        self.title["font"] = ("Calibri", "20", "bold")
        self.title.pack()

        # Nome do Fornecedor
        self.nomeFornecedor = Label(self.container2, text="Fornecedor ", font=self.fontePadrao)
        self.nomeFornecedor.pack(side=LEFT)

        self.inputFornecedor = Entry(self.container2,textvariable=nomeFornecedor)
        self.inputFornecedor["width"] = 30
        self.inputFornecedor["font"] = self.fontePadrao
        self.inputFornecedor.pack(side=LEFT)

        # Transportadora
        self.transportadora = Label(self.container2, text="Transportadora ", font=self.fontePadrao)
        self.transportadora.pack(side=LEFT)

        self.inputTransportadora = Entry(self.container2, textvariable=nomeTransportadora)
        self.inputTransportadora["width"] = 30
        self.inputTransportadora["font"] = self.fontePadrao
        self.inputTransportadora.pack(side=LEFT)

        # Nome do Motorista
        self.nomeMotorista = Label(self.container3, text="Nome do Motorista ", font=self.fontePadrao)
        self.nomeMotorista.pack(side=LEFT)

        self.inputMotorista = Entry(self.container3, textvariable=nomeMotorista)
        self.inputMotorista["width"] = 30
        self.inputMotorista["font"] = self.fontePadrao
        self.inputMotorista.pack(side=LEFT)

        # Nome do Produto
        self.produto = Label(self.container3, text="Produto ", font=self.fontePadrao)
        self.produto.pack(side=LEFT)

        self.inputProduto = Entry(self.container3, textvariable=nomeProduto)
        self.inputProduto["width"] = 30
        self.inputProduto["font"] = self.fontePadrao
        self.inputProduto.pack(side=LEFT)

        # Data e Hora de Chegada
        self.dataHoraChegada = Label(self.container4, text="Data/Hora de Chegada (DD/MM/AAAA HH:MM) ",
                                     font=self.fontePadrao)
        self.dataHoraChegada.pack(side=LEFT)

        self.inputDataHoraChegada = Entry(self.container4, textvariable=dataHoraChegada)
        self.inputDataHoraChegada["width"] = 20
        self.inputDataHoraChegada["font"] = self.fontePadrao
        self.inputDataHoraChegada.pack(side=LEFT)

        # Data e Hora de Saída
        self.dataHoraSaida = Label(self.container4, text="Data/Hora de Saída ", font=self.fontePadrao)
        self.dataHoraSaida.pack(side=LEFT)

        self.inputDataHoraSaida = Entry(self.container4, textvariable=dataHoraSaida)
        self.inputDataHoraSaida["width"] = 20
        self.inputDataHoraSaida["font"] = self.fontePadrao
        self.inputDataHoraSaida.pack(side=LEFT)

        # Número do CT-e
        self.numeroCTe = Label(self.container5, text="Número do CT-e ", font=self.fontePadrao)
        self.numeroCTe.pack(side=LEFT)

        self.inputCte = Entry(self.container5, textvariable=numeroCte)
        self.inputCte["width"] = 10
        self.inputCte["font"] = self.fontePadrao
        self.inputCte.pack(side=LEFT)

        # Número da NF
        self.nf = Label(self.container5, text="Número da NF ", font=self.fontePadrao)
        self.nf.pack(side=LEFT)

        self.inputNF = Entry(self.container5, textvariable=numeroNF)
        self.inputNF["width"] = 10
        self.inputNF["font"] = self.fontePadrao
        self.inputNF.pack(side=LEFT)

        # Peso da NF
        self.pesoNF = Label(self.container5, text="Peso da NF ", font=self.fontePadrao)
        self.pesoNF.pack(side=LEFT)

        self.inputPeso = Entry(self.container5, textvariable=pesoNF)
        self.inputPeso["width"] = 10
        self.inputPeso["font"] = self.fontePadrao
        self.inputPeso.pack(side=LEFT)

        # Motivo da Estadia
        self.motivoEstadia = Label(self.container6, text="Motivo da Estadia ", font=self.fontePadrao)
        self.motivoEstadia.pack(side=LEFT)

        self.inputMotivo = Entry(self.container6, textvariable=motivoEstadia)
        self.inputMotivo["width"] = 60
        self.inputMotivo["font"] = self.fontePadrao
        self.inputMotivo.pack(side=LEFT)

        # Button - Chama a função que limpa todos os campos
        self.btnInput = Button(self.container7, text="Novo", font=self.fontePadrao, width=20,
                               command=limparDados)
        self.btnInput.pack(side=LEFT)

        # Button - Chama a função que extrai os campos do PDF
        self.btnBuscar = Button(self.container7, text="Importar PDF", font=self.fontePadrao, width=20,
                                command=importarPdf)
        self.btnBuscar.pack(side=LEFT)

        # Button - Chama a função que salva dados do input
        self.btnInput = Button(self.container7, text="Emitir Estadia", font=self.fontePadrao, width=20,
                               command=preencherPlanilha)
        self.btnInput.pack(side=RIGHT)



root = Tk()
Application(root)
root.mainloop()
