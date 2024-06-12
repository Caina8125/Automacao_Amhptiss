import threading
import customtkinter
from tkinter import *
import tkinter.messagebox
from PIL import Image, ImageTk
from itertools import count, cycle
from Presentation.Desktop.Gif import ImageLabel
from Presentation.Desktop.SelecioneAutomacao import TelaSelecioneAutomacoes
import Application.AppService.AutomacaoService.AutenticacaoAppService as AuhtAppService

class Login:
    def __init__(self, janela,corTema):

        self.corJanela = corTema
        self.tela = janela

        self.barra = customtkinter.CTkLabel(self.tela ,text="", fg_color="#274360",width=650,height=40)
        self.barra.pack(padx=0, pady=0)
        
        self.photo = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\logo.png"), size=(80,90))
        self.botaoDark = customtkinter.CTkButton(self.tela,text="",image=self.photo, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark.pack()

        self.login = customtkinter.CTkLabel(self.tela, text="Fazer Login", text_color="#274360")
        self.login.pack(padx=3, pady=3)

        self.email = customtkinter.CTkEntry(self.tela,width=180,placeholder_text="Seu login do Amhptiss",placeholder_text_color="#274360",border_color="#274360")
        self.email.pack(padx=10, pady=10)

        self.senha = customtkinter.CTkEntry(self.tela, width=180,placeholder_text="Sua senha do Amhptiss",placeholder_text_color="#274360",border_color="#274360" ,show="*")
        self.senha.pack(padx=10, pady=10)

        self.comboBoxSetor = customtkinter.CTkComboBox(self.tela,border_color="#274360",button_color="#274360", button_hover_color="White",dropdown_fg_color="#274360", dropdown_text_color="White",text_color="#274360",values=["Glosa","Faturamento","Financeiro","Nota_Fiscal"], width=180)
        self.comboBoxSetor.pack(padx=10, pady=10)
        
        self.botaoLogar = customtkinter.CTkButton(self.tela, fg_color="#274360",width=130,text="Login", command=lambda: threading.Thread(target=self.clique).start())
        self.botaoLogar.pack(padx=5, pady=5)

    def modoEscuro(self):
        if (self.corJanela == "light"):
            self.corJanela = "dark"
            customtkinter.set_appearance_mode(self.corJanela)
            customtkinter.set_default_color_theme("dark-blue")
        else:
            self.corJanela = "light"
            customtkinter.set_appearance_mode(self.corJanela)

    def ocultarBotaoDark(self):
        self.botaoDark.pack_forget()

    def ocultarTelaLogin(self):
        self.login.pack_forget()
        self.email.pack_forget()
        self.senha.pack_forget()
        self.comboBoxSetor.pack_forget()
        self.botaoLogar.pack_forget()

    def reiniciarTelaLogin(self):
        ImageLabel.ocultarGif(self)
        self.login.pack(padx=10, pady=10)
        self.email.pack(padx=10, pady=10)
        self.senha.pack(padx=10, pady=10)
        self.comboBoxSetor.pack(padx=10, pady=10)
        self.botaoLogar.pack(padx=10, pady=10)

    def clique(self):
        setor = self.comboBoxSetor.get()
        loginUsuario = self.email.get()
        senhaUsuario = self.senha.get()

        self.ocultarTelaLogin()
        ImageLabel.iniciarGif(self,janela=self.tela,texto="Autenticando Usuário...")

        try:
            autenticacao = AuhtAppService.Autenticar(loginUsuario,senhaUsuario,setor)

            if (autenticacao[0]):
                ImageLabel.ocultarGif(self)
                self.ocultarBotaoDark()
                TelaSelecioneAutomacoes(setor, self.tela,token=autenticacao[1])
            else:
                tkinter.messagebox.showerror("Erro Autenticação", f"Usuário informado não tem permissão no setor {setor}")
                self.reiniciarTelaLogin()

        except Exception as e:
            tkinter.messagebox.showerror("Erro Autenticação", f"Usuário não autenticado! \nSeu login ou senha estão incorretos!")
            self.reiniciarTelaLogin()

    
        

        
