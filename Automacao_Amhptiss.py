from Presentation.Desktop.Main import Aplication
import os
import configparser
import semantic_version
import tkinter.messagebox
import asyncio

class Automacao:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read(r'Config\Config.ini')

        self.pathConfigAtualiza = self.config['ConfigPath']['PathAtualizaIni']
        self.pathArquivoConfigDev = self.config['ConfigPath']['pathConfigDiscoC']

        self.versaoA = self.versaoAtualiza()
        self.versaoC = self.versaoLocal()

        self.caminho = r"C:\Instaladores\InstalaAutomacao.exe"

        if self.versaoC < self.versaoA:
            try:
                os.execv(self.caminho, [self.caminho])
            except Exception as e:
                tkinter.messagebox.showerror("Erro", f"Erro ao chamar Executável da Instalação {e}")
        else:
            Aplication().janela.mainloop()

    def versaoLocal(self):
        self.configLocal = configparser.ConfigParser()
        self.configLocal.read(self.pathArquivoConfigDev)

        versaoC = semantic_version.Version(self.configLocal['ConfigVersion']['Version'])

        return versaoC

    def versaoAtualiza(self):
        self.configAtualiza = configparser.ConfigParser()
        self.configAtualiza.read(self.pathConfigAtualiza)

        versaoA = semantic_version.Version(self.configAtualiza['ConfigVersion']['Version'])

        return versaoA
        
Automacao()
        
