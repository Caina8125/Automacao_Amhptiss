import threading
import customtkinter
from tkinter import *
import tkinter.messagebox
from PIL import Image, ImageTk
from itertools import count, cycle
import Application.AppService.AutenticacaoAppService as AuhtAppService
import Application.AppService.ObterAutomacaoSetorAppService as AutomacoesAppService
import Application.AppService.ChamarAutomacoesAppService as chamarAutomacoes

class Application:
    def __init__(self):
        self.corAtual = "Custom"
        customtkinter.set_appearance_mode(self.corAtual)
        # customtkinter.set_default_color_theme("light-blue")

        self.janela = customtkinter.CTk()
        self.janela.geometry("550x400")
        self.janela.title("AMHP - Automações")
        # self.janela.configure(fg_color="White")
        self.janela.resizable(width=False, height=False)
        self.janela.eval('tk::PlaceWindow . center')
        self.janela.overrideredirect(False)

        self.barra = customtkinter.CTkLabel(self.janela,text="", bg_color="#274360")
        self.barra.pack(padx=340, pady=12)

        self.photo = PhotoImage(file = r"C:\Automacao_Amhptiss\Infra\Arquivos\logo.png", width=80)
        # self.logo =  customtkinter.CTkLabel(self.janela, text="",image=self.photo)
        # self.logo.pack(padx=10, pady=10)

        button_style = {
            "borderwidth": 0,  # Define a largura da borda como 0 para ocultá-la
            "highlightthickness": 0,  # Remove a espessura do destaque da borda
            "padx": 0,  # Define o preenchimento horizontal como 0
            "pady": 0,  # Define o preenchimento vertical como 0
        }
        self.botaoDark =  customtkinter.CTkButton(self.janela,text="",width=50, **button_style,image=self.photo, command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark.pack()

        self.login = customtkinter.CTkLabel(self.janela, text="Fazer Login")
        self.login.pack()

        self.email = customtkinter.CTkEntry(self.janela,width=180,placeholder_text="Seu login do Amhptiss")
        self.email.pack(padx=10, pady=10)

        self.senha = customtkinter.CTkEntry(self.janela, width=180,placeholder_text="Sua senha do Amhptiss", show="*")
        self.senha.pack(padx=10, pady=10)

        self.comboBoxSetor = customtkinter.CTkComboBox(self.janela,values=["GLOSA","FATURAMENTO","FINANCEIRO","TESOURARIA"], width=140)
        self.comboBoxSetor.pack(padx=10, pady=10)

        
        self.botaoLogar = customtkinter.CTkButton(self.janela, width=80,fg_color="#274360",text="Login", command=lambda: threading.Thread(target=self.clique).start())
        # self.botaoLogar.configure(ba)
        self.botaoLogar.pack(padx=10, pady=10)

    def modoEscuro(self):
        if (self.corAtual == "light"):
            self.corAtual = "dark"
            customtkinter.set_appearance_mode(self.corAtual)
            customtkinter.set_default_color_theme("dark-blue")
        else:
            self.corAtual = "light"
            customtkinter.set_appearance_mode(self.corAtual)

    def clique(self):
        setor = self.comboBoxSetor.get()
        loginUsuario = self.email.get()
        senhaUsuario = self.senha.get()

        self.ocultar()
        # self.chamarGift()

        try:
            autenticacao = AuhtAppService.Autenticar(loginUsuario,senhaUsuario,setor)

            if (autenticacao):
                self.ocultar()
                self.telaPrincipal(setor)

                # tkinter.messagebox.showinfo("Autenticado", f"Usuário {loginUsuario} Autenticado com sucesso! \nPertence ao setor {setor}")
                # return True
            else:
                tkinter.messagebox.showerror("Erro Autenticação", f"Usuário informado não tem permissão no setor {setor}")
                self.reiniciarTelaLogin()

        except:
            tkinter.messagebox.showerror("Erro Autenticação", f"Usuário não autenticado! \nSeu login ou senha estão incorretos!")
            self.reiniciarTelaLogin()

    def ocultar(self):
        self.login.pack_forget()
        self.email.pack_forget()
        self.senha.pack_forget()
        self.comboBoxSetor.pack_forget()
        self.botaoLogar.pack_forget()

    def chamarGift(self):
        self.LabelGift = customtkinter.CTkLabel(self.janela, text="Trabalhando")
        self.LabelGift.pack(padx=10, pady=10)
        # self.lbl = ImageLabel(self.LabelGift)
        self.lbl.load(r"C:\Automacao_Amhptiss\Infra\Arquivos\loader2.gif")

    def reiniciarTelaLogin(self):
        self.login.pack(padx=10, pady=10)
        self.email.pack(padx=10, pady=10)
        self.senha.pack(padx=10, pady=10)
        self.comboBoxSetor.pack(padx=10, pady=10)
        self.botaoLogar.pack(padx=10, pady=10)

    def telaPrincipal(self,setorUsuario):
        self.comboAutomacoes(setorUsuario)
        self.botaoIniciarAutomacao()

    def comboAutomacoes(self,setorUsuario):
        self.listaAutomacao = AutomacoesAppService.obterListaSetor(setorUsuario)
        self.comboBoxAutomacao = customtkinter.CTkComboBox(self.janela,values=self.listaAutomacao, width=300)
        self.comboBoxAutomacao.pack(padx=10, pady=25)

    def botaoIniciarAutomacao(self):
        self.iniciarAutomacao = customtkinter.CTkButton(self.janela, width=80,fg_color="#274360",text="Iniciar", command=lambda: threading.Thread(target=self.chamarAutomacao).start())
        self.iniciarAutomacao.pack(padx=10, pady=10)

    def chamarAutomacao(self):
        automacaoSelecionada = self.comboBoxAutomacao.get()
        chamarAutomacoes.aplicarChamada(automacaoSelecionada)
        

        
