import customtkinter
import Infra.Repository.AutenticacaoRepository

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

        self.checkbox = customtkinter.CTkCheckBox(self.janela, text="Lembrar Login")
        self.checkbox.pack(padx=10, pady=10)

        self.botao = customtkinter.CTkButton(self.janela, text="Login", command=clique)
        self.botao.pack(padx=10, pady=10)

        def clique():
            autenticar = Infra.Repository.AutenticacaoRepository(self.login, self.senha)

        

Application().janela.mainloop()
