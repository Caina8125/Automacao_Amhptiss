import os
import shutil
import customtkinter
import threading
import tkinter.messagebox

class Atualiza:
    def __init__(self):
        customtkinter.set_appearance_mode("dark")
        self.janela = customtkinter.CTk()
        self.janela.geometry("400x100")
        self.janela.title("AMHP - Automações")
        self.janela.resizable(width=False, height=False)
        self.janela.eval('tk::PlaceWindow . center')
        self.janela.overrideredirect(True)
        
        # self.janela.overrideredirect(True)
        self.pathLocal = r"C:\Automacao_Amhptiss"
        self.exeLocal = "Automacao_Amhptiss.exe"
        self.pathAtualiza = r"\\10.0.0.239\atualiza\Automacao_Amhptiss"
        self.pathConfigLocal = r"C:\Automacao_Amhptiss\Config\Config.ini"
        self.pathConfigAtualiza = r"\\10.0.0.239\atualiza\Automacao_Amhptiss\Config\Config.ini"

        self.verificaInstalacao = self.fazerVerificacao()

        if (self.verificaInstalacao):
            self.textoInfoA = customtkinter.CTkLabel(self.janela, text="Existe uma atualização a ser feita na Automação")
            self.textoInfoA.pack(padx=10, pady=10)
            self.buttonAtualiza = customtkinter.CTkButton(self.janela, fg_color="#274360",width=130,text="Atualizar", command=lambda: threading.Thread(target=self.Atualizar).start())
            self.buttonAtualiza.pack(padx=10, pady=10)
        else:
            self.textoInfo = customtkinter.CTkLabel(self.janela, text="Clique em instalar para iniciar.")
            self.textoInfo.pack(padx=10, pady=10)
            self.buttonInstala = customtkinter.CTkButton(self.janela, fg_color="#274360",width=130,text="Instalar", command=lambda: threading.Thread(target=self.Instalar).start())
            self.buttonInstala.pack(padx=10, pady=10)

    def label(self,texto):
        self.textoAguarde = customtkinter.CTkLabel(self.janela, text=texto, font=("Arial",20))
        self.textoAguarde.pack(padx=3, pady=23)
        self.janela.update_idletasks()

    def chamarAutomacao(self):
        self.caminho = r"C:\Automacao_Amhptiss\Automacao_Amhptiss.exe"
        try:
            os.execv(self.caminho, [self.caminho])
        except Exception as e:
            tkinter.messagebox.showerror("Erro", f"Erro ao chamar Executável da Instalação {e}")

    def fazerVerificacao(self):
        if not os.path.isdir(self.pathLocal):
            return False
        else:
            return True
        
    def Instalar(self):
        self.textoInfo.pack_forget()
        self.buttonInstala.pack_forget()
        if not os.path.isdir(self.pathLocal):
            try:
                self.label("Aguarde! \nInstalando Automação...")
                os.makedirs(self.pathLocal)
                shutil.rmtree(self.pathLocal)  # Remove a pasta de destino se ela já existir
                shutil.copytree(self.pathAtualiza, self.pathLocal)
                self.chamarAutomacao()
            except Exception as e:
                tkinter.messagebox.showerror("Erro", f"Erro ao tentar Excluir pasta {e}")
                self.janela.after(500,lambda:self.janela.destroy())
    
    def Atualizar(self):
        self.textoInfoA.pack_forget()
        self.buttonAtualiza.pack_forget()
        if os.path.isdir(self.pathLocal):
            self.label("Aguarde! \nAtualizando Automação...")
            # self.janela.update_idletasks
            try:
                # shutil.rmtree(self.pathLocal)
                self.remove_directory_except(self.pathLocal, 'Output')
                shutil.copytree(self.pathAtualiza, self.pathLocal)
                self.chamarAutomacao()
            except Exception as e:
                try:
                    shutil.copytree(self.pathAtualiza, self.pathLocal,dirs_exist_ok=True)
                    self.chamarAutomacao()
                except Exception as e:
                    tkinter.messagebox.showerror("Erro", f"Erro ao tentar Copiar pasta com dirs_exist_ok {e}")
                    self.janela.after(500,lambda:self.janela.destroy())

    def remove_directory_except(self, directory, except_folder):
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            # Verifica se o item é o diretório que queremos manter
            if os.path.basename(item_path) == except_folder:
                continue
            # Remove arquivos e diretórios
            if os.path.isdir(item_path):
                shutil.rmtree(item_path)
            else:
                os.remove(item_path)


Atualiza().janela.mainloop()