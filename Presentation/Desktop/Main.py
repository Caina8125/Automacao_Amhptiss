import customtkinter
from Presentation.Desktop.Login import Login
from Presentation.Desktop.TesteAwait import chamada_assincrona
import tkinter as tk

class Aplication(tk.Tk):
    def __init__(self):
        self.corAtual = "light"
        customtkinter.set_appearance_mode(self.corAtual)
        self.janela = customtkinter.CTk()
        window_width = 640
        window_height = 480
        self.center_window(window_width, window_height)
        # self.janela.geometry("640x480")
        
        self.janela.title("AMHP - Automações")
        self.janela.iconbitmap(r"Infra\Arquivos\Robo.ico")
        self.janela.resizable(width=False, height=False)
        # self.janela.eval('tk::PlaceWindow . center')
        Login(self.janela, self.corAtual)
        # chamada_assincrona(self)

    def modoEscuro(self):
        if (self.corAtual == "light"):
            self.corAtual = "dark"
            customtkinter.set_appearance_mode(self.corAtual)
            customtkinter.set_default_color_theme("dark-blue")
        else:
            self.corAtual = "light"
            customtkinter.set_appearance_mode(self.corAtual)

    def center_window(self, width, height):
        # Obter as dimensões da tela
        screen_width = self.janela.winfo_screenwidth()
        screen_height = self.janela.winfo_screenheight()
        
        # Calcular a posição do ponto superior esquerdo para centralizar a janela
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        
        # Definir a geometria da janela
        self.janela.geometry(f'{width}x{height}+{x}+{y}')