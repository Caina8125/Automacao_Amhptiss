import threading
from tkinter.filedialog import askopenfilename
import customtkinter
import tkinter as tk
from tkinter import *
from PIL import Image
from Presentation.Desktop.Gif import ImageLabel
from Presentation.Desktop.Selecao import Selecao

class SelecaoUmArquivo(Selecao):
    def __init__(self, janela, obj):
        super().__init__(janela)

        self.diretorioInput = customtkinter.CTkEntry(self.terceiroContainer, "Selecione um arquivo", width=420, border_color="#274360", state="readonly")
        self.diretorioInput.pack(padx=0, side=LEFT)

        self.btnSelecionarDiretorio = customtkinter.CTkButton(self.terceiroContainer, text="...", width=30, height=28, fg_color="#274360", command=self.selecionar_arquivo)
        self.btnSelecionarDiretorio.pack(side=LEFT, padx=1)
        
        self.iniciar = customtkinter.CTkButton(self.quartoContainer, width=80,fg_color="#274360",text="Iniciar", command=lambda: threading.Thread(target=self.iniciar_automacao(obj)).start())
        self.iniciar.pack(pady=0, padx=0)

    def selecionar_arquivo(self):
        file_path = askopenfilename()
        if file_path:
            self.arquivoInput.configure(state=NORMAL)
            self.arquivoInput.delete(0, tk.END)
            self.arquivoInput.insert(0, file_path)
            self.arquivoInput.configure(state='readonly')

    def iniciar_automacao(self, obj):
        self.ocultar_widgets()
        ImageLabel.iniciarGif(self,janela=self.terceiroContainer,texto="Trabalhando...")
        arquivo = self.arquivoInput.get()
        threading.Thread(target=lambda: self.run_function(obj, arquivo)).start()

    def run_function(self, obj, arquivo):
        obj.inicia_automacao(arquivo)
        ImageLabel.ocultarGif(self)