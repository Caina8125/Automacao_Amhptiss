import threading
import customtkinter
import tkinter as tk
from tkinter import *
from tkinter import ttk
import tkinter.messagebox
from PIL import Image, ImageTk
from Presentation.Desktop.Gif import ImageLabel
import Application.AppService.FaturasPreFaturadasAppService as Faturas
import Application.AppService.ObterCaminhoFaturasAppService as ObterCaminho

class telaTabelaFaturas:
    def __init__(self,janela,token,nomeAutomacao,codigoConvenio):

        self.tela = janela
        self.corJanela = "light"

        self.photo = customtkinter.CTkImage(light_image = Image.open(r"C:\Automacao_Amhptiss\Infra\Arquivos\logo.png"), size=(60,70))
        self.botaoDark = customtkinter.CTkButton(self.tela,text="",image=self.photo, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark.pack()
        
        self.listas = Faturas.obterListaFaturas("normal",codigoConvenio,400,"2023/01/01","2024/05/30",token)
        self.treeView(listas=self.listas)
        self.botaoBuscarFaturas.pack(padx=10, pady=10)
        
    def sort_treeview(self,tree, col, reverse):
        # Função para ordenar a treeview com base no cabeçalho clicado
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        data.sort(reverse=reverse)
        for index, item in enumerate(data):
            tree.move(item[1], '', index)
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))

    def treeView(self, listas):
        

        self.style = ttk.Style()
        self.style.configure("Treeview",
                             foreground="White",
                             background='#274360',
                             rowheight=23)
        
        self.style.map('Treeview',
                background=[('selected','gray')],
                foreground=[('selected','black')])
        
        self.qtd = len(listas) 
        self.texto = customtkinter.CTkLabel(self.tela, text=f"{self.qtd} Faturas Pré-Faturadas", text_color="#274360",font=("Arial",20))
        self.texto.pack(padx=3, pady=3)

        self.my_tree = ttk.Treeview(self.tela)

        self.my_tree['columns'] = ("Faturas","Remessas","Convenio","Usuario","Status","Arquivo")

        self.my_tree.column("#0",width=0, stretch=NO)
        self.my_tree.column("Faturas", anchor=CENTER, width=70)
        self.my_tree.column("Remessas",anchor=CENTER, width=70)
        self.my_tree.column("Convenio",anchor=CENTER, width=100)
        self.my_tree.column("Usuario", anchor=CENTER, width=140)
        self.my_tree.column("Status", anchor=CENTER, width=140)
        self.my_tree.column("Arquivo", anchor=CENTER, width=100)

        for col in ("Faturas","Remessas","Convenio","Usuario","Status","Arquivo"):
            self.my_tree.heading(col, text=col,anchor=CENTER,command=lambda c=col: self.sort_treeview(self.my_tree, c, False))

        i = 0
        for lista in listas:
            self.my_tree.insert(parent='', index='end', values=lista)
            i =+ 1
        self.my_tree.pack(padx=3, pady=2)
        
        self.botaoStart = customtkinter.CTkButton(self.tela, fg_color="#274360",width=80,text="Enviar", command=lambda: threading.Thread(target=self.iniciar()).start())
        # self.botaoStart.place(x=178, y=340)

        self.botaoRemover = customtkinter.CTkButton(self.tela, fg_color="#b30505",width=80,text="Remover",command=lambda: threading.Thread(target=self.remover()).start())
        # self.botaoRemover.place(x=280, y=340)

        self.botaoRemoverFaturasEnviadas = customtkinter.CTkButton(self.tela, fg_color="#058288",width=80,text="Remover Faturas Enviadas", command=lambda: threading.Thread(target=self.remover()).start())
        # self.botaoRemover.place(x=280, y=340)

        self.botaoBuscarFaturas = customtkinter.CTkButton(self.tela, width=80,text="Buscar Faturas Escaneadas ", command=lambda: threading.Thread(target=self.buscarFaturas()).start())
        # self.botaoBuscarFaturas.pack(padx=10, pady=10)
    

    def remover(self):
        itenSelecionado = self.my_tree.selection()
        try:
            for iten in itenSelecionado:
                self.my_tree.delete(iten)
        except:
            tkinter.messagebox.showinfo("Erro", f"Selecione uma fatura para remover")

    def buscarFaturas(self):
        self.ocultarTreeView()
        self.faturas = []
        for item in self.my_tree.get_children():
            self.valores = self.my_tree.item(item, 'values')
            self.faturas.append(self.valores)

        listaCaminhoFaturas = ObterCaminho.IniciarBusca(self.faturas)
        listaCaminhoFaturas.sort_values('Status Envio', ascending=False, inplace=True)
        df = listaCaminhoFaturas.values.tolist()
        self.reiniciarTreeView(listaAtualizada=df)

    def reiniciarTreeView(self,listaAtualizada):
        # self.botaoRemoverFaturasEnviadas.place(x=318, y=420)
        self.botaoStart.place(x=220, y=420)
        self.botaoRemover.place(x=330, y=420)
        self.treeView(listaAtualizada)

    def ocultarTreeView(self):
        self.texto.pack_forget()
        self.my_tree.pack_forget()
        self.botaoStart.place_forget()
        self.botaoRemover.place_forget()
        self.botaoBuscarFaturas.pack_forget()

    def modoEscuro(self):
        if (self.corJanela == "light"):
            self.corJanela = "dark"
            customtkinter.set_appearance_mode(self.corJanela)
            customtkinter.set_default_color_theme("dark-blue")
        else:
            self.corJanela = "light"
            customtkinter.set_appearance_mode(self.corJanela)