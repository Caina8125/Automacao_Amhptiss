import customtkinter
import Application.AppService.AutenticacaoAppService as AuhtAppService
import threading
import tkinter.messagebox

class Application:
    def __init__(self, master=None):
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("dark-blue")

        self.janela = customtkinter.CTk()
        self.janela.geometry("500x300")
        
        self.login = customtkinter.CTkLabel(self.janela, text="Fazer Login")
        self.login.pack(padx=10, pady=10)

        self.email = customtkinter.CTkEntry(self.janela, placeholder_text="Seu e-mail")
        self.email.pack(padx=10, pady=10)

        self.senha = customtkinter.CTkEntry(self.janela, placeholder_text="Sua senha", show="*")
        self.senha.pack(padx=10, pady=10)

        self.comboBox = customtkinter.CTkComboBox(self.janela,values=["GLOSA","FATURAMENTO","FINANCEIRO","TESOURARIA"], width=140)
        self.comboBox.pack()

        self.botao = customtkinter.CTkButton(self.janela, text="Login", command=lambda: threading.Thread(target=self.clique).start())
        self.botao.pack(padx=10, pady=10)

    def clique(self):
        setor = self.comboBox.get()
        loginUsuario = self.email.get()
        senhaUsuario = self.senha.get()
        try:
            autenticacao = AuhtAppService.Autenticar(loginUsuario,senhaUsuario,setor)
            if (autenticacao):
                pass
            else:
                tkinter.messagebox.showerror("Erro Autenticação", f"Usuário informado não tem permissão no setor {setor}")
        except:
            tkinter.messagebox.showerror("Erro Autenticação", f"Usuário não autenticado! \nSeu login ou senha estão incorretos!")
            




        
