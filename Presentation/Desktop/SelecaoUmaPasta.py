import threading
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showwarning
import customtkinter
import tkinter as tk
from tkinter import *
from Presentation.Desktop.Gif import ImageLabel
from Presentation.Desktop.Selecao import Selecao

class SelecaoUmaPasta(Selecao):
    def __init__(self, janela, obj):
        super().__init__(janela)

        self.diretorioInput = customtkinter.CTkEntry(self.terceiroContainer, width=420, border_color="#274360", state="readonly")
        self.diretorioInput.pack(padx=0, side=LEFT)

        self.btnSelecionarDiretorio = customtkinter.CTkButton(self.terceiroContainer, text="...", width=30, height=28, fg_color="#274360", command=self.selecionarDiretorio)
        self.btnSelecionarDiretorio.pack(side=LEFT, padx=1)
        
        self.iniciar = customtkinter.CTkButton(self.quartoContainer, width=80,fg_color="#274360",text="Iniciar", command=lambda: threading.Thread(target=self.iniciar_automacao(obj)).start())
        self.iniciar.pack(pady=0, padx=0)

    def selecionarDiretorio(self):
        folder_path = askdirectory()
        if folder_path:
            self.diretorioInput.configure(state=NORMAL)
            self.diretorioInput.delete(0, tk.END)
            self.diretorioInput.insert(0, folder_path)
            self.diretorioInput.configure(state='readonly')

    def iniciar_automacao(self, obj):
        if self.diretorioInput.get() == '':
            showwarning(message='Selecione uma pasta!')
            return
        
        self.ocultar_widgets()
        ImageLabel.iniciarGif(self,janela=self.terceiroContainer,texto="Trabalhando...")
        diretorio = self.diretorioInput.get()
        threading.Thread(target=lambda: self.run_function(obj, diretorio)).start()

    def run_function(self, obj, dir):
        obj.inicia_automacao(diretorio=dir)
        ImageLabel.ocultarGif(self)
        