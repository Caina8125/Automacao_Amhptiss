import customtkinter
from Presentation.Desktop.Login import Login

class Aplication:
    def __init__(self):
        self.corAtual = "light"
        customtkinter.set_appearance_mode(self.corAtual)
        self.janela = customtkinter.CTk()
        self.janela.geometry("550x400")
        self.janela.title("AMHP - Automações")
        self.janela.iconbitmap(r"C:\Automacao_Amhptiss\Infra\Arquivos\Robo.ico")
        self.janela.resizable(width=False, height=False)
        self.janela.eval('tk::PlaceWindow . center')
        Login(self.janela, self.corAtual)

    def modoEscuro(self):
        if (self.corAtual == "light"):
            self.corAtual = "dark"
            customtkinter.set_appearance_mode(self.corAtual)
            customtkinter.set_default_color_theme("dark-blue")
        else:
            self.corAtual = "light"
            customtkinter.set_appearance_mode(self.corAtual)

autenticao = Aplication().janela.mainloop()