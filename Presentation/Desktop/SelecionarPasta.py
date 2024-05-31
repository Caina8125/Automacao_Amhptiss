import threading
from tkinter.filedialog import askdirectory
import customtkinter
import tkinter as tk
from tkinter import *
from tkinter import ttk
import tkinter.messagebox
from PIL import Image, ImageTk
from Presentation.Desktop.Gif import ImageLabel
import Application.AppService as service

class TelaSelecionarPasta:
    def __init__(self, janela, obj):
        super().__init__()
        self.tela = janela

        self.primeiroContainer = Frame(self.tela)
        self.primeiroContainer.pack()

        self.modoEscuroAut()

        self.segundoContainer = Frame(self.tela)
        self.segundoContainer.pack()

        self.terceiroContainer = Frame(self.tela)
        self.terceiroContainer.pack(pady=40)

        self.quartoContainer = Frame(self.tela)
        self.quartoContainer.pack()

        self.informePastaLbl = customtkinter.CTkLabel(self.tela, text="Selecione uma pasta", text_color="#274360", bg_color="transparent")
        self.informePastaLbl.place(x=93, y=147)

        self.diretorioInput = customtkinter.CTkEntry(self.terceiroContainer, width=420, border_color="#274360", state="readonly")
        self.diretorioInput.pack(padx=0, side=LEFT)

        self.btnSelecionarDiretorio = customtkinter.CTkButton(self.terceiroContainer, text="...", width=30, height=28, fg_color="#274360", command=self.selecionarDiretorio)
        self.btnSelecionarDiretorio.pack(side=LEFT, padx=1)
        
        self.iniciar = customtkinter.CTkButton(self.quartoContainer, width=80,fg_color="#274360",text="Iniciar", command=lambda: threading.Thread(target=self.iniciar_automacao(obj)).start())
        self.iniciar.pack(pady=0, padx=0)

    def modoEscuroAut(self):
        self.photo = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\logo.png"), size=(80,90))
        self.botaoDark = customtkinter.CTkButton(self.primeiroContainer ,text="",image=self.photo, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark.pack()

    def selecionarDiretorio(self):
        folder_path = askdirectory()
        if folder_path:
            self.diretorioInput.configure(state=NORMAL)
            self.diretorioInput.delete(0, tk.END)
            self.diretorioInput.insert(0, folder_path)
            self.diretorioInput.configure(state='readonly')

    def iniciar_automacao(self, obj):
        self.ocultar_widgets()
        ImageLabel.iniciarGif(self,janela=self.terceiroContainer,texto="Trabalhando...")
        diretorio = self.diretorioInput.get()
        threading.Thread(target=lambda: self.run_function(obj, diretorio)).start()

    def run_function(self, obj, diretorio):
        obj.inicia_automacao(diretorio)
        ImageLabel.ocultarGif(self)
    
    def ocultar_widgets(self):
        self.informePastaLbl.place_forget()
        self.diretorioInput.pack_forget()
        self.iniciar.pack_forget()
        self.btnSelecionarDiretorio.pack_forget()
        