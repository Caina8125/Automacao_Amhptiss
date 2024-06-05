from abc import ABC, abstractmethod
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

class Selecao(ABC):
    def __init__(self, janela):
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

        self.quintoContainer = Frame(self.tela)
        self.quintoContainer.pack()

    def modoEscuroAut(self):
        self.photo = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\logo.png"), size=(80,90))
        self.botaoDark = customtkinter.CTkButton(self.primeiroContainer ,text="",image=self.photo, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark.pack()

    @abstractmethod
    def iniciar_automacao(self, obj):...
    
    def ocultar_widgets(self):
        self.diretorioInput.pack_forget()
        self.iniciar.pack_forget()
        self.btnSelecionarDiretorio.pack_forget()
        