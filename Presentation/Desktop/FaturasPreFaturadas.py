import os
import threading
from time import sleep
import customtkinter
from tkinter import *
from tkinter import ttk
import tkinter.messagebox
from PIL import Image
from pandas import DataFrame
from Presentation.Desktop.Gif import ImageLabel
import Application.AppService.FaturasPreFaturadasAppService as Faturas
import Application.AppService.GuiasFaturaAppService as Guias
import Application.AppService.ObterCaminhoFaturasAppService as ObterCaminho

class telaTabelaFaturas():
    def __init__(self, janela,token,obj,codigoConvenio, setor):
        super().__init__()
        self.tela = janela
        self.token = token
        self.setorUsuario = setor
        self.obj = obj
        self.exibirBotaoDark()
        ImageLabel.iniciarGif(self,janela=self.tela,texto="Buscando faturas no \nAMHPTISS...")
        
        self.codigoConvenio = codigoConvenio
        self.token = token
        self.corJanela = "light"

        threading.Thread(target=self.pegar_dados).start()
        
    def pegar_dados(self):
        self.listas = Faturas.obterListaFaturas("normal",self.codigoConvenio,400,"2023/01/01","2024/05/30",self.token)
        print(self.listas)
        self.show_data()
        # self.after(0, self.show_data, self.listas)

    def show_data(self):
        ImageLabel.ocultarGif(self)
        # self.ocultarBotaoDark()
        # img = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\voltar.png"), size=(20,30))

        # self.exibirBotaoDark()
        self.treeView(listas=self.listas)
        self.botaoBuscarFaturas.pack(padx=10, pady=10)
        # self.botaoVoltar = customtkinter.CTkButton(self.tela, hover=False, fg_color='transparent', bg_color="transparent",width=80,text="", image=img, command=lambda: threading.Thread(target=self.botao_voltar_click).start())
        # self.botaoVoltar.place(y=40, x=20)

        
        # Fecha a tela de carregamento após mostrar os dados por um tempo
        # self.after(0, self.destroy)

    def sort_treeview(self,tree, col, reverse):
        # Função para ordenar a treeview com base no cabeçalho clicado
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        data.sort(reverse=reverse)
        for index, item in enumerate(data):
            tree.move(item[1], '', index)
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))

    def treeView(self, listas):
        self.qtd = len(listas) 
        self.texto = customtkinter.CTkLabel(self.tela, text=f"{self.qtd} Faturas Pré-Faturadas", text_color="#274360",font=("Arial",20))
        self.texto.pack(padx=3, pady=3)

        self.style = ttk.Style()
        self.style.configure("Treeview",
                             foreground="White",
                             background='#274360',
                             rowheight=23)
        
        self.style.map('Treeview',
                background=[('selected','gray')],
                foreground=[('selected','black')])
        
        self.my_tree = ttk.Treeview(self.tela)
        self.my_tree['columns'] = ("Faturas", "Remessas", "Convênio", "Usuário", "Protocolos", "Status", "Arquivo")
        self.my_tree.column("#0",width=0, stretch=NO)
        self.my_tree.column("Faturas", anchor=CENTER, width=70)
        self.my_tree.column("Remessas",anchor=CENTER, width=70)
        self.my_tree.column("Convênio",anchor=CENTER, width=100)
        self.my_tree.column("Usuário", anchor=CENTER, width=140)
        self.my_tree.column("Protocolos", anchor=CENTER, width=70)
        self.my_tree.column("Status", anchor=CENTER, width=140)
        self.my_tree.column("Arquivo", anchor=CENTER, width=100)

        for col in ("Faturas", "Remessas", "Convênio", "Usuário", "Protocolos", "Status", "Arquivo"):
            self.my_tree.heading(col, text=col,anchor=CENTER,command=lambda c=col: self.sort_treeview(self.my_tree, c, False))

        i = 0
        for lista in listas:
            self.my_tree.insert(parent='', index='end', values=lista)
            i =+ 1
        self.my_tree["displaycolumns"]=("Faturas", "Remessas", "Convênio", "Usuário", "Status", "Arquivo")
        self.my_tree.pack(padx=3, pady=2)
        
        self.botaoStart = customtkinter.CTkButton(self.tela, fg_color="#274360",width=80,text="Enviar", command=lambda: threading.Thread(target=self.iniciar, args=(self.obj,)).start())
        # self.botaoStart.place(x=178, y=340)

        self.botaoRemover = customtkinter.CTkButton(self.tela, fg_color="#b30505",width=80,text="Remover",command=lambda: threading.Thread(target=self.remover).start())
        # self.botaoRemover.place(x=280, y=340)

        self.botaoRemoverFaturasEnviadas = customtkinter.CTkButton(self.tela, fg_color="#058288",width=80,text="Remover Faturas Enviadas", command=lambda: threading.Thread(target=self.remover()).start())
        # self.botaoRemover.place(x=280, y=340)

        self.botaoBuscarFaturas = customtkinter.CTkButton(self.tela, width=80,text="Buscar Faturas Escaneadas ", command=lambda: threading.Thread(target=self.buscarFaturas()).start())
        # self.botaoBuscarFaturas.pack(padx=10, pady=10)

    def get_treeview_data(self):
        full_treeview = []

        for row in self.my_tree.get_children():
            full_treeview.append(self.my_tree.item(row)["values"])
        return full_treeview
    
    def iniciar(self, obj):
        self.ocultarTreeView()
        ImageLabel.iniciarGif(self,janela=self.tela,texto="Enviando Suas Faturas Escaneadas \nno Portal...")
        threading.Thread(target=self.exec_automacao, args=(obj,)).start()

    def exec_automacao(self, obj):
        faturas = self.obterFaturasEncontradas()
        df = DataFrame(faturas)
        df.columns = ['Fatura', 'Remessa', 'Convênio', 'Usuário', 'Protocolo', 'Status Fatura', 'Nome Arquivo', 'Caminho']

        if self.codigoConvenio == 225:
            dados_faturas = self.get_guias_fatura(df)
            obj.inicia_automacao(dados_faturas=dados_faturas, df_treeview=...)

        else:
            obj.inicia_automacao(dados_fatura=faturas)

        ImageLabel.ocultarGif(self)

    def botaoRoboPaz(self):
        self.photoPaz = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\RoboEnvia.png"), size=(100,100))
        self.botaoDark2 = customtkinter.CTkButton(self.tela,text="",image=self.photoPaz, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark2.pack()

    def get_guias_fatura(self, df: DataFrame):
        dict_cols = {j: i for i, j in enumerate(df.columns)}
        lista = {"faturas": []}

        for row in df.values:
            fatura = row[dict_cols['Fatura']]
            caminho = row[dict_cols['Caminho']]
            protocolo = row[dict_cols['Protocolo']]
            lista_guias = Guias.obterListaGuias("normal", 225, fatura, self.token)

            dados_guia = []

            for guia in lista_guias:
                caminho_guia = f"{caminho}\\{guia}.pdf"

                if os.path.isfile(caminho_guia):
                    dados_guia.append(
                        {
                            "guia": guia,
                            "caminho_guia": caminho_guia
                        }
                    )

                else:
                    dados_guia.append(
                        {
                            "guia": guia,
                            "caminho_guia": caminho_guia
                        }
                    )

            lista["faturas"].append(
                    {
                        "fatura": fatura,
                        "protocolo": protocolo,
                        "lista_guias": [
                            *dados_guia
                        ]
                    }
                )

        return lista

    def obterFaturasTela(self):
        self.faturas = []

        for item in self.my_tree.get_children():
            valores_tupla = self.my_tree.item(item, 'values')
            valores_lista = [*valores_tupla]
            del valores_lista[-2:]
            self.faturas.append(valores_lista)

        return self.faturas

    def obterFaturasEncontradas(self):
        self.faturas = []
        for item in self.my_tree.get_children():
            self.valores = self.my_tree.item(item, 'values')
            self.status = self.valores[5]
            if self.status != "Não Encontrado":
                self.faturas.append(self.valores)

        return self.faturas

    def remover(self):
        itenSelecionado = self.my_tree.selection()
        try:
            for iten in itenSelecionado:
                self.my_tree.delete(iten)
        except:
            tkinter.messagebox.showinfo("Erro", f"Selecione uma fatura para remover")

    def buscarFaturas(self):
        self.ocultarTreeView()
        ImageLabel.iniciarGif(self,janela=self.tela,texto="Buscando faturas\nescaneadas...")
        threading.Thread(target=self.exec_busca_faturas_escaneadsas).start()
        
    def exec_busca_faturas_escaneadsas(self):
        sleep(2)
        self.listaTela = self.obterFaturasTela()
        listaCaminhoFaturas = ObterCaminho.IniciarBusca(self.listaTela, self.codigoConvenio)
        listaCaminhoFaturas.sort_values('Status Envio', ascending=False, inplace=True)
        df = listaCaminhoFaturas.values.tolist()
        ImageLabel.ocultarGif(self)
        self.reiniciarTreeView(listaAtualizada=df)    

    def reiniciarTreeView(self,listaAtualizada):
        # self.botaoRemoverFaturasEnviadas.place(x=318, y=420)
        # self.botaoDark.pack()
        self.botaoStart.place(x=220, y=420)
        self.botaoRemover.place(x=330, y=420)
        self.treeView(listaAtualizada)

    def ocultarTreeView(self):
        self.botaoStart.place_forget()
        self.botaoRemover.place_forget()
        # self.botaoDark.pack_forget()
        self.texto.pack_forget()
        self.my_tree.pack_forget()
        self.botaoBuscarFaturas.pack_forget()

    def ocultarBotoes(self):
        self.botaoStart.destroy()
        self.botaoRemover.destroy()
    
    def ocultarBotaoDark(self):
        self.botaoDark.pack_forget()

    def modoEscuro(self):
        if (self.corJanela == "light"):
            self.corJanela = "dark"
            customtkinter.set_appearance_mode(self.corJanela)
            customtkinter.set_default_color_theme("dark-blue")
        else:
            self.corJanela = "light"
            customtkinter.set_appearance_mode(self.corJanela)

    def exibirBotaoDark(self):
        self.photo = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\logo.png"), size=(60,70))
        self.botaoDark = customtkinter.CTkButton(self.tela,text="",image=self.photo, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark.pack()

    # def botao_voltar_click(self):
    #     self.botaoDark.pack_forget()
    #     self.texto.pack_forget()
    #     self.ocultarTreeView()
    #     TelaSelecioneAutomacoes(self.setorUsuario, self.tela, self.token)
        
        try:
            self.botaoBuscarFaturas.pack_forget()
        except:
            try:
                self.botaoRemover.place_forget()
                self.botaoStart.place_forget()
            except:
                pass