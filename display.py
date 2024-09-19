from tkinter import *
from tkinter import filedialog, messagebox, ttk
import pdfplumber
from openpyxl.reader.excel import load_workbook


class Application():
    def __init__(self, master=None):

        def dataDeSaida(text):
            data = text.split('\n')[4].split(' ')[2].replace('.', '/')
            hora = text.split('\n')[4].split(' ')[3].split(':')[0]
            minutos = text.split('\n')[4].split(' ')[3].split(':')[1]

            return data + ' ' + hora + ':' + minutos

        def escolherPdf():
            self.pdf = filedialog.askopenfile(initialdir="/Desktop", title="Selecione um arquivo",
                                              filetypes=[("Arquivos PDF", "*.pdf")])
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

            self.inputFornecedor.focus()  # Puxa o foco para o campo Fornecedor

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
            campoMotivos = str(
                self.inputMotivo.get().upper())  # Motivo, vai se iniciar com "MOTIVO: " e concatenar com o real motivo da estadia

            campos = [campoFornecedor, campoTransportadora, campoMotorista, campoProduto, campoEntrada, campoSaida,
                      campoCte, campoNumeroNF, campoPesoNF, campoMotivos]

            fields = [self.inputFornecedor, self.inputTransportadora, self.inputMotorista, self.inputProduto,
                      self.inputDataHoraChegada, self.inputDataHoraSaida, self.inputCte, self.inputNF, self.inputPeso,
                      self.inputMotivo]

            for i in range(10):
                if campos[i] in '':
                    mensagemErro = "Preencha todos os campos!"
                    messagebox.showinfo('Aviso!', mensagemErro)
                    fields[i].focus()
                    return

            preencherPlanilha()

        def preencherPlanilha():
            wb = load_workbook('estadia\Cálculo estadia.xlsx')  # Carrega o arquivo existente
            planilha = wb.active  # Seleciona a planilha ativa

            # Pega os dados dos inputs e coloca na planilha de acordo com a célula referenciada
            planilha['F3'] = str(self.inputTransportadora.get().upper())  # Transportador
            planilha['C4'] = int(self.inputNF.get())  # NF
            planilha['C5'] = str(self.inputProduto.get().upper())  # Produto, impor condição se for Rocha!!
            planilha['F12'] = str(self.inputPeso.get())  # Peso NF com vírgula
            planilha['B15'] = str(self.inputDataHoraSaida.get())  # Data e Hora da saída no formato DD/MM/AA HH:MM
            planilha['B9'] = str(self.inputDataHoraChegada.get())  # Data e Hora de chegada no formato DD/MM/AA HH:MM
            planilha['C3'] = str(self.inputFornecedor.get().upper())  # Fornecedor
            planilha['F4'] = str(self.inputCte.get())  # Ct-e
            planilha['F5'] = str(self.inputMotorista.get().upper())  # Motorista
            planilha['E16'] = 'MOTIVO: ' + str(
                self.inputMotivo.get().upper())  # Motivo, vai se iniciar com "MOTIVO: " e concatenar com o real motivo da estadia

            salvarPlanilha(wb)

        def salvarPlanilha(wb):
            self.planilha = filedialog.asksaveasfilename(
                initialdir="C:/Users/adrie/PycharmProjects/pythonProject1/EstadiasCalculadas", defaultextension=".xlsx",
                title="Salvar como", filetypes=[("Excel", "*.xlsx")])

            if self.planilha:
                wb.save(self.planilha)
                wb.close()

        # -----------------------------------------------------------------------------------------

        # PARTE GRÁFICA

        def containers():
            self.fontePadrao = ("Arial", "10")

            self.container1 = Frame(self.root)
            self.container1["pady"] = 10
            self.container1.pack()

            self.container2 = Frame(self.root)
            self.container2["padx"] = 20
            self.container2["pady"] = 5
            self.container2.pack()

            self.container3 = Frame(self.root)
            self.container3["padx"] = 20
            self.container3["pady"] = 5
            self.container3.pack()

            self.container4 = Frame(self.root)
            self.container4["padx"] = 20
            self.container4["pady"] = 5
            self.container4.pack()

            self.container5 = Frame(self.root)
            self.container5["padx"] = 20
            self.container5["pady"] = 5
            self.container5.pack()

            self.container6 = Frame(self.root)
            self.container6["padx"] = 20
            self.container6["pady"] = 5
            self.container6.pack()

            self.container7 = Frame(self.root)
            self.container7["padx"] = 20
            self.container7["pady"] = 5
            self.container7.pack()

            labels(self.container1, self.container2, self.container3, self.container4, self.container5, self.container6,
                   self.fontePadrao)
            buttons(self.container7, self.fontePadrao)

        def labels(container1, container2, container3, container4, container5, container6, fontePadrao):
            # Titulo
            self.title = Label(container1, text="Calculo de Estadia")
            self.title["font"] = ("Calibri", "20", "bold")
            self.title.pack()

            # Nome do Fornecedor
            self.nomeFornecedor = Label(container2, text="Fornecedor ", font=fontePadrao)
            self.nomeFornecedor.pack(side=LEFT)

            self.inputFornecedor = Entry(container2, textvariable=nomeFornecedor, width=30, font=fontePadrao)
            self.inputFornecedor.focus()
            self.inputFornecedor.pack(side=LEFT)

            # Transportadora
            self.transportadora = Label(container2, text="Transportadora ", font=fontePadrao)
            self.transportadora.pack(side=LEFT)

            transportadorasCadastradas = ['FUTURO', 'G10', 'CARVALHO', 'FRIBON', 'D GRANEL', 'SIMOES BEBEDOURO', 'AGUETONI']
            trasnportadorasCadastradas.sort()

            self.inputTransportadora = ttk.Combobox(container2, textvariable=nomeTransportadora, values=trasnportadorasCadastradas, width=30, font=fontePadrao)
            self.inputTransportadora.pack(side=LEFT)

            # Nome do Motorista
            self.nomeMotorista = Label(container3, text="Nome do Motorista ", font=fontePadrao)
            self.nomeMotorista.pack(side=LEFT)

            self.inputMotorista = Entry(container3, textvariable=nomeMotorista, width=30, font=fontePadrao)
            self.inputMotorista.pack(side=LEFT)

            # Nome do Produto
            self.produto = Label(container3, text="Produto ", font=fontePadrao)
            self.produto.pack(side=LEFT)

            produtosCadastrados = ['ROCHA UMA', 'ROCHA CMISS', 'MICRO R P', 'KCL 00-00-58 F','KCL 00-00-58 GR','SSP 00-19-00','TSP 00-46-00 GR', 'MAP 11-52-00 F', 'KCL 00-00-60 GR IMP', 'CAL','MICRO HMoNi', 'ENXOFRE F IMP']
            produtosCadastrados.sort()
            self.inputProduto = ttk.Combobox(container3, textvariable=nomeProduto, values=produtodosCadastrados, width=30, font=fontePadrao)
            self.inputProduto.pack(side=LEFT)

            # Data e Hora de Chegada
            self.dataHoraChegada = Label(container4, text="Data/Hora de Chegada ",font=fontePadrao)
            self.dataHoraChegada.pack(side=LEFT)

            self.inputDataHoraChegada = Entry(container4, textvariable=dataHoraChegada, width=20, font=fontePadrao)
            self.inputDataHoraChegada.pack(side=LEFT)

            # Data e Hora de Saída
            self.dataHoraSaida = Label(container4, text="Data/Hora de Saída ", font=fontePadrao)
            self.dataHoraSaida.pack(side=LEFT)

            self.inputDataHoraSaida = Entry(container4, textvariable=dataHoraSaida, width=20, font=fontePadrao)
            self.inputDataHoraSaida.pack(side=LEFT)

            # Número do CT-e
            self.numeroCTe = Label(container5, text="Número do CT-e ", font=fontePadrao)
            self.numeroCTe.pack(side=LEFT)

            self.inputCte = Entry(container5, textvariable=numeroCte, width=10, font=fontePadrao)
            self.inputCte.pack(side=LEFT)

            # Número da NF
            self.nf = Label(container5, text="Número da NF ", font=fontePadrao)
            self.nf.pack(side=LEFT)

            self.inputNF = Entry(container5, textvariable=numeroNF, width=10, font=fontePadrao)
            self.inputNF.pack(side=LEFT)

            # Peso da NF
            self.pesoNF = Label(container5, text="Peso da NF ", font=fontePadrao)
            self.pesoNF.pack(side=LEFT)

            self.inputPeso = Entry(container5, textvariable=pesoNF, width=10, font=fontePadrao)
            self.inputPeso.pack(side=LEFT)

            # Motivo da Estadia
            self.motivoEstadia = Label(container6, text="Motivo da Estadia ", font=fontePadrao)
            self.motivoEstadia.pack(side=LEFT)

            self.inputMotivo = Entry(container6, textvariable=motivoEstadia, width=60, font=fontePadrao)
            self.inputMotivo.pack(side=LEFT)

        def buttons(container7, fontePadrao):

            # Button - Chama a função que limpa todos os campos
            self.btnInput = Button(container7, text="Novo", font=fontePadrao, width=20,
                                   command=limparDados)
            self.btnInput.pack(side=LEFT)

            # Button - Chama a função que extrai os campos do PDF
            self.btnBuscar = Button(container7, text="Importar PDF", font=fontePadrao, width=20,
                                    command=importarPdf)
            self.btnBuscar.pack(side=LEFT)

            # Button - Chama a função que salva dados do input
            self.btnInput = Button(container7, text="Emitir Estadia", font=fontePadrao, width=20,
                                   command=conferenciaDeDados)
            self.btnInput.pack(side=RIGHT)
            self.btnInput.place()

        def janela():
            self.root.title('Cálculo de Estadia')
            containers()

        self.root = root

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

        janela()


root = Tk()
Application(root)
root.mainloop()
