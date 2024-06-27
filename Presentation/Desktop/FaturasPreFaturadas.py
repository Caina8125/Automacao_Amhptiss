from abc import ABC, abstractmethod
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os
import threading
from time import sleep
import customtkinter
from tkinter import *
from tkinter import ttk
import tkinter.messagebox
from PIL import Image
from pandas import DataFrame
import customtkinter as ctk
from Presentation.Desktop.Gif import ImageLabel
import Application.AppService.FaturasPreFaturadasAppService as Faturas
import Application.AppService.GuiasFaturaAppService as Guias
import Application.AppService.ObterCaminhoFaturasAppService as ObterCaminho

class telaTabelaFaturas(ABC):
    def tela_tabela_faturas(self, janela,token,obj,codigoConvenio, setor):
        super().__init__()
        self.data_atual = datetime.now()
        self.seis_meses_anterior = self.data_atual - relativedelta(months=6)
        self.tela = janela
        self.container1 = ctk.CTkFrame(self.tela, bg_color='transparent', fg_color='transparent')
        self.container1.pack()
        self.container2 = ctk.CTkFrame(self.tela, bg_color='transparent', fg_color='transparent')
        self.container2.pack()
        self.container3 = ctk.CTkFrame(self.tela, bg_color='transparent', fg_color='transparent')
        self.container3.pack()
        self.container4 = ctk.CTkFrame(self.tela, bg_color='transparent', fg_color='transparent')
        self.container4.pack()
        self.container5 = ctk.CTkFrame(self.tela, bg_color='transparent', fg_color='transparent')
        self.container5.pack()
        self.token = token
        self.setorUsuario = setor
        self.obj = obj
        self.exibirBotaoDark()
        ImageLabel.iniciarGif(self,janela=self.container2,texto="Buscando faturas no \nAMHPTISS...")
        
        self.codigoConvenio = codigoConvenio
        self.token = token
        self.corJanela = "light"

        threading.Thread(target=self.pegar_dados).start()
        
    def pegar_dados(self):
        data_atual_str = self.data_atual.strftime('%Y/%m/%d')
        seis_meses_anterior_str = self.seis_meses_anterior.strftime('%Y/%m/%d')
        self.listas = Faturas.obterListaFaturas("normal",self.codigoConvenio,400, seis_meses_anterior_str, data_atual_str,self.token)
        self.show_data()
        # self.after(0, self.show_data, self.listas)

    def show_data(self):
        ImageLabel.ocultarGif(self)
        # self.ocultarBotaoDark()
        img = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\voltar.png"), size=(20,30))

        # self.exibirBotaoDark()
        self.treeView(listas=self.listas)
        self.botaoBuscarFaturas.pack(padx=10, pady=10)
        self.botaoVoltar = customtkinter.CTkButton(self.tela, hover=False, fg_color='#274360', bg_color="#274360", width=80, text="", image=img, command=self.botao_voltar_click)
        self.botaoVoltar.place(y=1, x=1)

        
        # Fecha a tela de carregamento após mostrar os dados por um tempo
        # self.after(0, self.destroy)

    def toggle_selection(self):
        all_items = self.my_tree.get_children()
        selected_items = self.my_tree.selection()
        
        if len(all_items) == len(selected_items):
            self.my_tree.selection_remove(all_items)
        else:
            self.my_tree.selection_set(all_items)

    def sort_treeview(self,tree, col, reverse):
        # Função para ordenar a treeview com base no cabeçalho clicado
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        data.sort(reverse=reverse)
        for index, item in enumerate(data):
            tree.move(item[1], '', index)
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))

    def treeView(self, listas):
        self.qtd = len(listas) 
        self.texto = customtkinter.CTkLabel(self.container2, text=f"{self.qtd} Faturas Pré-Faturadas", text_color="#274360",font=("Arial",20))
        self.texto.pack(padx=3, pady=3)

        self.style = ttk.Style()
        self.style.configure("Treeview",
                             foreground="White",
                             background='#274360',
                             rowheight=23)
        
        self.style.map('Treeview',
                background=[('selected','gray')],
                foreground=[('selected','black')])
        
        self.my_tree = ttk.Treeview(self.container3)
        self.my_tree['columns'] = ("Faturas", "Remessas", "Convênio", "Usuário", "Protocolos", "Envio Operadora", "Status", "Arquivo")
        self.my_tree.column("#0",width=0, stretch=NO)
        self.my_tree.column("Faturas", anchor=CENTER, width=70)
        self.my_tree.column("Remessas",anchor=CENTER, width=70)
        self.my_tree.column("Convênio",anchor=CENTER, width=100)
        self.my_tree.column("Usuário", anchor=CENTER, width=140)
        self.my_tree.column("Protocolos", anchor=CENTER, width=70)
        self.my_tree.column("Envio Operadora", anchor=CENTER, width=70)
        self.my_tree.column("Status", anchor=CENTER, width=140)
        self.my_tree.column("Arquivo", anchor=CENTER, width=100)

        for col in ("Faturas", "Remessas", "Convênio", "Usuário", "Protocolos", "Envio Operadora", "Status", "Arquivo"):
            self.my_tree.heading(col, text=col,anchor=CENTER,command=lambda c=col: self.sort_treeview(self.my_tree, c, False))

        i = 0
        for lista in listas:
            self.my_tree.insert(parent='', index='end', values=lista)
            i =+ 1
        self.my_tree["displaycolumns"]=("Faturas", "Remessas", "Convênio", "Usuário", "Status", "Arquivo")
        self.scrollbar = ctk.CTkScrollbar(self.container3, orientation='vertical', command=self.my_tree.yview)
        self.my_tree.configure(yscroll=self.scrollbar.set)
        self.my_tree.pack(padx=3, pady=2, side=LEFT)
        self.scrollbar.pack(side=RIGHT, fill='y')
        
        # self.botaoRemover.place(x=280, y=340)

        self.botaoRemoverFaturasEnviadas = customtkinter.CTkButton(self.container4, fg_color="#058288",width=80,text="Remover Faturas Enviadas", command=lambda: threading.Thread(target=self.remover()).start())
        # self.botaoRemover.place(x=280, y=340)

        self.botaoBuscarFaturas = customtkinter.CTkButton(self.container4, width=80,text="Buscar Faturas Escaneadas ", command=lambda: threading.Thread(target=self.buscarFaturas).start())
        self.botaoRemoverFaturasEnviadas = ctk.CTkButton(self.container4, fg_color="#058288",width=80,text="Remover Faturas Enviadas", command=lambda: threading.Thread(target=self.remover()).start())
        # self.botaoBuscarFaturas.pack(padx=10, pady=10)

    def get_treeview_data(self):
        full_treeview = []

        for row in self.my_tree.get_children():
            full_treeview.append(self.my_tree.item(row)["values"])
        return full_treeview
    
    def iniciar(self, obj):
        faturas = self.obterFaturasEncontradas()
        if not faturas:
            tkinter.messagebox.showwarning('', 'Não há faturas para enviar!')
            return
        self.ocultarTreeView()
        self.botaoSelectAll.place_forget()
        self.botaoVoltar.configure(state="disabled")
        ImageLabel.iniciarGif(self,janela=self.container2,texto="Enviando Suas Faturas Escaneadas \nno Portal...")
        threading.Thread(target=self.exec_automacao, args=(obj,)).start()

    def exec_automacao(self, obj):
        faturas = self.obterFaturasEncontradas()
        df = DataFrame(faturas)
        colunas_df = ['Fatura', 'Remessa', 'Convênio', 'Usuário', 'Protocolo', 'Envio Operadora', 'Status Fatura', 'Nome Arquivo', 'Caminho']
        df.columns = colunas_df
        
        if self.codigoConvenio == 160:
            ImageLabel.atualizaLabel(self, 'Obtendo números da nota fiscal...')
            df_treeview = self.get_nf_fatura(DataFrame(self.get_treeview_data()))
            colunas_df.append('Nota Fiscal')
            ImageLabel.atualizaLabel(self, 'Enviando no portal...')
        else:
            df_treeview = DataFrame(self.get_treeview_data())

        df_treeview.columns = colunas_df

        if self.codigoConvenio == 225:
            ImageLabel.atualizaLabel(self, 'Obtendo as guias das faturas...')
            dados_faturas = self.get_guias_fatura(df)
            ImageLabel.atualizaLabel(self, 'Enviando no portal...')
            df_treeview = obj.inicia_automacao(dados_faturas=dados_faturas, df_treeview=df_treeview, token=self.token)

        else:
            df_treeview = obj.inicia_automacao(df_treeview=df_treeview, token=self.token)

        ImageLabel.ocultarGif(self)
        self.reiniciarTreeView(df_treeview.values.tolist())
        self.botaoSelectAll.place(y=103, x=0)
        self.botaoVoltar.configure(state="normal")

    def botaoRoboPaz(self):
        self.photoPaz = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\RoboEnvia.png"), size=(100,100))
        self.botaoDark2 = customtkinter.CTkButton(self.tela,text="",image=self.photoPaz, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark2.pack()

    def get_nf_fatura(self, df_treeview: DataFrame):
        df_treeview['Nota Fiscal'] = ''
        faturas_list = df_treeview[0].values.tolist()

        for fatura in faturas_list:
            n_nota_fiscal = Guias.obterNotaFiscalFatura("normal", self.codigoConvenio, fatura, self.token)
            if isinstance(n_nota_fiscal, str):
                n_nota_fiscal = n_nota_fiscal.split('/')[1]
                df_treeview.loc[df_treeview[0] == fatura, 'Nota Fiscal'] = n_nota_fiscal
    
        return df_treeview

    def get_guias_fatura(self, df: DataFrame):
        dict_cols = {j: i for i, j in enumerate(df.columns)}
        lista = {"faturas": []}

        for row in df.values:
            fatura = row[dict_cols['Fatura']]
            caminho = row[dict_cols['Caminho']]
            protocolo = row[dict_cols['Protocolo']]
            lista_guias = Guias.obterListaGuias("normal", self.codigoConvenio, fatura, self.token)

            dados_guia = []

            for guia in lista_guias:
                caminho_guia = f"{caminho}\\{guia}.pdf"

                if os.path.isfile(caminho_guia):
                    dados_guia.append(
                        {
                            "guia": guia,
                            "caminho_guia": caminho_guia,
                            "guia_enviada": False
                        }
                    )

                else:
                    dados_guia.append(
                        {
                            "guia": guia,
                            "caminho_guia": caminho_guia,
                            "guia_enviada": False
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
            self.status = self.valores[6]
            if self.status != "Fatura Não Encontrada" and self.status != "Enviada":
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
        self.scrollbar.pack_forget()
        self.botaoVoltar.configure(state="disabled")
        ImageLabel.iniciarGif(self,janela=self.container2,texto="Buscando faturas\nescaneadas...")
        threading.Thread(target=self.exec_busca_faturas_escaneadsas).start()
        
    def exec_busca_faturas_escaneadsas(self):
        sleep(1)
        self.listaTela = self.obterFaturasTela()
        listaCaminhoFaturas = ObterCaminho.IniciarBusca(self.listaTela, self.codigoConvenio)
        listaCaminhoFaturas.sort_values('Status Envio', ascending=False, inplace=True)
        df = listaCaminhoFaturas.values.tolist()
        ImageLabel.ocultarGif(self)
        self.reiniciarTreeView(listaAtualizada=df)
        img = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\select_all.png"), size=(30,40))
        self.botaoSelectAll = customtkinter.CTkButton(self.tela, hover=False, fg_color='transparent', bg_color="transparent",width=80,text="", image=img, command=self.toggle_selection)
        self.botaoSelectAll.place(y=103, x=0)
        self.botaoVoltar.configure(state="normal")    

    def reiniciarTreeView(self,listaAtualizada):
        # self.botaoRemoverFaturasEnviadas.place(x=318, y=420)
        # self.botaoDark.pack()
        self.treeView(listaAtualizada)
        self.botaoStart = customtkinter.CTkButton(self.container4, fg_color="#274360",width=80,text="Enviar", command=lambda: threading.Thread(target=self.iniciar, args=(self.obj,)).start())
        self.botaoRemover = customtkinter.CTkButton(self.container4, fg_color="#b30505",width=80,text="Remover",command=lambda: threading.Thread(target=self.remover).start())
        self.botaoStart.pack(side=LEFT, pady=20, padx=30)
        self.botaoRemover.pack(side=LEFT, pady=20, padx=30)

    def ocultarTreeView(self):
        try:
            self.botaoStart.pack_forget()
            self.botaoRemover.pack_forget()
        except:
            pass
        self.scrollbar.pack_forget()
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
        self.botaoDark = customtkinter.CTkButton(self.container1,text="",image=self.photo, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark.pack()

    @abstractmethod
    def botao_voltar_click(self):
        ...
        # self.botaoDark.pack_forget()
        # self.texto.pack_forget()
        # self.ocultarTreeView()
        # TelaSelecioneAutomacoes(self.setorUsuario, self.tela, self.token)
        
        # try:
        #     self.botaoBuscarFaturas.pack_forget()
        # except:
        #     try:
        #         self.botaoRemover.place_forget()
        #         self.botaoStart.place_forget()
        #     except:
        #         pass