from abc import ABC
from datetime import datetime
import tkinter
from tkinter.filedialog import askopenfilenames
from tkinter.messagebox import showwarning
from dateutil.relativedelta import relativedelta
from tkinter import *
from tkinter import ttk
import customtkinter as ctk
from PIL import Image
import threading

from pandas import DataFrame

from Presentation.Desktop.Gif import ImageLabel
import Application.AppService.FaturasPreFaturadasAppService as Faturas

class tela_busca_protocolo(ABC):
    def tela_busca_protocolo(self, janela,token,obj,codigoConvenio, setor):
        super().__init__()
        self.data_atual = datetime.now()
        self.seis_meses_anterior = self.data_atual - relativedelta(months=6)
        self.obj = obj
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

        threading.Thread(target=self.pegar_dados_busca).start()

    def pegar_dados_busca(self):
        data_atual_str = self.data_atual.strftime('%Y/%m/%d')
        seis_meses_anterior_str = self.seis_meses_anterior.strftime('%Y/%m/%d')
        self.listas = Faturas.get_faturas_data("normal",self.codigoConvenio,400, seis_meses_anterior_str, data_atual_str,self.token)
        self.show_data_busca()
    
    def show_data_busca(self):
        ImageLabel.ocultarGif(self)
        # self.ocultarBotaoDark()
        img_voltar = ctk.CTkImage(light_image = Image.open(r"Infra\Arquivos\voltar.png"), size=(20,30))
        img_select_all = ctk.CTkImage(light_image = Image.open(r"Infra\Arquivos\select_all.png"), size=(30,40))

        # self.exibirBotaoDark()
        self.treeViewBusca(listas=self.listas)
        self.botaoConferirEnvio.pack(side=LEFT, pady=20, padx=30)
        self.botaoRemover.pack(side=LEFT, pady=20, padx=30)
        self.botaoVoltar = ctk.CTkButton(self.tela, hover=False, fg_color='#274360', bg_color="#274360", width=80, text="", image=img_voltar, command=self.botao_voltar_click)
        self.botaoVoltar.place(y=1, x=1)
        self.botaoSelectAll = ctk.CTkButton(self.tela, hover=False, fg_color='transparent', bg_color="transparent",width=80,text="", image=img_select_all, command=self.toggle_selection)
        self.botaoSelectAll.place(y=103, x=0)

    def treeViewBusca(self, listas):
        self.qtd = len(listas) 
        self.texto = ctk.CTkLabel(self.container2, text=f"{self.qtd} Faturas Pré-Faturadas", text_color="#274360",font=("Arial",20))
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
        self.my_tree['columns'] = ("Fatura", "Protocolo", "Remessa", "Quantidade", "Valor Total", "Status")
        self.my_tree.column("#0",width=0, stretch=NO)
        self.my_tree.column("Fatura", anchor=CENTER, width=70)
        self.my_tree.column("Protocolo",anchor=CENTER, width=70)
        self.my_tree.column("Remessa",anchor=CENTER, width=100)
        self.my_tree.column("Quantidade", anchor=CENTER, width=100)
        self.my_tree.column("Valor Total", anchor=CENTER, width=100)
        self.my_tree.column("Status", anchor=CENTER, width=140)

        for col in ("Fatura", "Protocolo", "Remessa", "Quantidade", "Valor Total", "Status"):
            self.my_tree.heading(col, text=col,anchor=CENTER,command=lambda c=col: self.sort_treeview(self.my_tree, c, False))

        i = 0
        for lista in listas:
            self.my_tree.insert(parent='', index='end', values=lista)
            i =+ 1
        self.my_tree["displaycolumns"]=("Fatura", "Protocolo", "Remessa", "Quantidade", "Valor Total", "Status")
        self.scrollbar = ctk.CTkScrollbar(self.container3, orientation='vertical', command=self.my_tree.yview)
        self.my_tree.configure(yscroll=self.scrollbar.set)
        self.my_tree.pack(padx=3, pady=2, side=LEFT)
        self.scrollbar.pack(side=RIGHT, fill='y')
        
        # self.botaoRemover.place(x=280, y=340)

        self.botaoRemover = ctk.CTkButton(self.container4, fg_color="#b30505",width=80,text="Remover",command=lambda: threading.Thread(target=self.remover_busca).start())
        # self.botaoRemover.place(x=280, y=340)

        self.botaoConferirEnvio = ctk.CTkButton(self.container4, width=80,text="Conferir Envio", command=lambda: threading.Thread(target=self.conferir_envio).start())
    
    def get_treeview_data(self):
        full_treeview = []

        for row in self.my_tree.get_children():
            full_treeview.append(self.my_tree.item(row)["values"])
        return full_treeview

    def conferir_envio(self):
        df_treeview = DataFrame(self.get_treeview_data())
        if df_treeview.empty:
            tkinter.messagebox.showwarning('', 'Não há faturas para conferir!')
            return
        self.ocultarTreeViewBuscar()
        self.botaoRemover.place_forget()
        self.botaoSelectAll.place_forget()
        self.botaoConferirEnvio.place_forget()
        self.botaoVoltar.configure(state="disabled")
        ImageLabel.iniciarGif(self,janela=self.container2,texto="Conferindo envios...")
        threading.Thread(target=self.exec_automacao_busca, args=(df_treeview,)).start()

    def exec_automacao_busca(self, df_treeview):
        df_treeview.columns = ["Fatura", "Protocolo", "Remessa", "Quantidade", "Valor Total", "Status"]
        showwarning('', 'Selecione Protocolos de Recebimentos!')
        files_path_list = askopenfilenames()
        df_treeview = self.obj.inicia_automacao(df_treeview=df_treeview, files_path_list=files_path_list)
        ImageLabel.ocultarGif(self)
        self.reiniciarTreeViewBusca(df_treeview.values.tolist())
        self.botaoVoltar.configure(state="normal")

    def reiniciarTreeViewBusca(self,listaAtualizada):
        # self.botaoRemoverFaturasEnviadas.place(x=318, y=420)
        # self.botaoDark.pack()
        self.treeViewBusca(listaAtualizada)
        self.botaoConferirEnvio.pack(side=LEFT, pady=20, padx=30)
        self.botaoRemover.pack(side=LEFT, pady=20, padx=30)
        
    def ocultarTreeViewBuscar(self):
        try:
            self.botaoConferirEnvio.pack_forget()
            self.botaoRemover.pack_forget()
        except:
            pass
        self.scrollbar.pack_forget()
        self.texto.pack_forget()
        self.my_tree.pack_forget()

    def remover_busca(self):
        itenSelecionado = self.my_tree.selection()
        try:
            for iten in itenSelecionado:
                self.my_tree.delete(iten)
        except:
            tkinter.messagebox.showinfo("Erro", f"Selecione uma fatura para remover")

    def toggle_selection(self):
        all_items = self.my_tree.get_children()
        selected_items = self.my_tree.selection()
        
        if len(all_items) == len(selected_items):
            self.my_tree.selection_remove(all_items)
        else:
            self.my_tree.selection_set(all_items)