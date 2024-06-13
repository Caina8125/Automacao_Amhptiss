from abc import ABC, abstractmethod
import threading
import customtkinter
from PIL import Image, ImageTk
from Presentation.Desktop.Gif import ImageLabel
import Application.AppService.ChamarTabelaAppService as Tabela
import Application.AppService.ObterAutomacaoSetorAppService as AutomacoesAppService
from Presentation.Desktop.FaturasPreFaturadas import telaTabelaFaturas
from Application import AppService
import Application.AppService.AutomacaoService as automacoes
import time
import tkinter as tk

from Presentation.Desktop.SelecaoUmaPasta import SelecaoUmaPasta


class TelaSelecioneAutomacoes(ABC):
    def iniciar_tela_selecione(self, setorUsuario, janela,token):
        super().__init__()
        self.setorUsuario = setorUsuario
        self.tela = janela
        # self.bottonSair(self.tela)
        self.modoEscuroAut()
        self.labelSelecioneAutomacao()
        self.comboAutomacoes(setorUsuario)
        self.botaoAvancar(token)
        
    def logOf(self):
        # self.bottonSair.pack_forget()
        self.ocultarTelaInicial()
        self.reiniciarTelaLogin()

    def modoEscuroAut(self):
        self.photo = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\logo.png"), size=(80,90))
        self.botaoDark = customtkinter.CTkButton(self.tela,text="",image=self.photo, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark.pack()

    def ocultarBotaoDark(self):
        self.botaoDark.pack_forget()

    def ocultarTelaInicial(self):
        self.iniciarAvancar.pack_forget()
        self.comboBoxAutomacao.pack_forget()
        self.labelAutomacao.pack_forget()

    def labelSelecioneAutomacao(self):
        self.labelAutomacao = customtkinter.CTkLabel(self.tela, text="Selecione uma Automação:", text_color="#274360")
        self.labelAutomacao.pack(padx=0, pady=0)

    def comboAutomacoes(self,setorUsuario):
        self.listaAutomacao = AutomacoesAppService.obterListaSetor(setorUsuario)
        self.comboBoxAutomacao = customtkinter.CTkComboBox(self.tela,values=self.listaAutomacao, border_color="#274360",button_color="#274360", button_hover_color="White",dropdown_fg_color="#274360", dropdown_text_color="White",text_color="#274360",width=300)
        self.comboBoxAutomacao.pack(padx=10, pady=25)

    def botaoAvancar(self,token):
        self.iniciarAvancar = customtkinter.CTkButton(self.tela, width=80,fg_color="#274360",text="Avançar", command=lambda: threading.Thread(target=self.avancar(token)).start())
        self.iniciarAvancar.pack(padx=10, pady=10)
    
    def ocultaBotaoAvancar(self):
        self.iniciarAvancar.pack_forget()

    @abstractmethod
    def avancar(self,token):
        ...
    #     # ImageLabel.reiniciarGif(self,janela=self.tela,texto="Buscando faturas no \nAMHPTISS...")
    #     # self.ocultaBotaoAvancar()
    #     # time.sleep(5)
    #     self.ocultarTelaInicial()
    #     automacaoSelecionada = self.comboBoxAutomacao.get()
    #     self.ocultarBotaoDark()

    #     match automacaoSelecionada:
    #         case "Faturamento - Anexar Guia Geap":
    #             telaTabelaFaturas(janela=self.tela,token=token,obj=automacoes.Geap('https://www2.geap.com.br/auth/prestadorVue.asp', '66661692120', 'Amhp2024'),codigoConvenio=225, setor=self.setorUsuario)

    #         case "Faturamento - Conferência de Protocolos":
    #             ...

    #         case "Faturamento - Conferência GEAP":
    #             ...

    #         case "Faturamento - Conferência Bacen":
    #             ...

    #         case "Faturamento - Enviar PDF Bacen":
    #             ...

    #         case "Faturamento - Enviar PDF Benner":
    #             ...

    #         case "Faturamento - Enviar PDF BRB":
    #             telaTabelaFaturas(janela=self.tela,token=token,obj=automacoes.EnviarPdf(),codigoConvenio=10, setor=self.setorUsuario)

    #         case "Faturamento - Enviar PDF GDF":
    #             telaTabelaFaturas(janela=self.tela,token=token,nomeAutomacao=automacaoSelecionada,codigoConvenio=433, setor=self.setorUsuario)

    #         case "Faturamento - Enviar XML Bacen":
    #             ...

    #         case "Faturamento - Enviar XML Benner":
    #             ...

    #         case "Faturamento - Enviar XML Caixa":
    #             ...

    #         case "Faturamento - Leitor de PDF GAMA":
    #             ...

    #         case "Faturamento - Verificar Situação BRB":
    #             ...

    #         case "Faturamento - Verificar Situação Fascal":
    #             ...

    #         case "Faturamento - Verificar Situação Gama":
    #             ...

    #         case "Financeiro - Buscar Faturas GEAP":
    #             ...

    #         case "Financeiro - Demonstrativos Amil":
    #             ...

    #         case "Financeiro - Demonstrativos BRB":
    #             ...

    #         case "Financeiro - Demonstrativos Câmara dos Deputados":
    #             ...

    #         case "Financeiro - Demonstrativos Camed":
    #             ...

    #         case "Financeiro - Demonstrativos Casembrapa":
    #             ...

    #         case "Financeiro - Demonstrativos Cassi":
    #             ...

    #         case "Financeiro - Demonstrativos Codevasf":
    #             ...

    #         case "Financeiro - Demonstrativos E-Vida":
    #             ...

    #         case "Financeiro - Demonstrativos Fapes":
    #             ...

    #         case "Financeiro - Demonstrativos Fascal":
    #             ...

    #         case "Financeiro - Demonstrativos Gama":
    #             ...

    #         case "Financeiro - Demonstrativos Life Empresarial":
    #             ...

    #         case "Financeiro - Demonstrativos MPU":
    #             ...

    #         case "Financeiro - Demonstrativos PMDF":
    #             ...

    #         case "Financeiro - Demonstrativos Postal":
    #             ...

    #         case "Financeiro - Demonstrativos Real Grandeza":
    #             ...

    #         case "Financeiro - Demonstrativos Saúde Caixa":
    #             ...

    #         case "Financeiro - Demonstrativos Serpro":
    #             ...

    #         case "Financeiro - Demonstrativos SIS":
    #             ...

    #         case "Financeiro - Demonstrativos STF":
    #             ...

    #         case "Financeiro - Demonstrativos TJDFT":
    #             ...

    #         case "Financeiro - Demonstrativos Unafisco":
    #             ...

    #         case "Glosa - Gerador de Planilha GDF":
    #             ...

    #         case "Glosa - Gerar Planilhas SERPRO":
    #             ...

    #         case "Glosa - Filtro Matrículas":
    #             ...

    #         case "Glosa - Recursar Amil":
    #             ...  

    #         case "Glosa - Recursar Benner(Câmara, CAMED, FAPES, Postal)":
    #             ...

    #         case "Glosa - Recursar BRB":
                
    #             SelecaoUmaPasta(self.tela, AppService.RecursoBrb('https://portal.saudebrb.com.br/GuiasTISS/Logon', '00735860000173', 'AMHP7356!'))

    #         case "Glosa - Recursar Casembrapa":
    #             ...

    #         case "Glosa - Recursar Cassi":
    #             ...

    #         case "Glosa - Recursar GEAP Duplicado":
    #             ...

    #         case "Glosa - Recursar GEAP Sem Duplicado":
    #             ...
            
    #         case "Glosa - Recursar E-VIDA":
    #             ...

    #         case "Glosa - Recursar Fascal":
    #             ...

    #         case "Glosa - Recursar Gama":
    #             ...

    #         case "Glosa - Recursar Petrobras":
    #             ...

    #         case "Glosa - Recursar Real Grandeza":
    #             ...

    #         case "Glosa - Recursar Saúde Caixa":
    #             ...

    #         case "Glosa - Recursar SIS":
    #             ...

    #         case "Glosa - Recursar STF":
    #             ...

    #         case "Glosa - Recursar STM":
    #             ...

    #         case "Glosa - Recursar TJDFT":
    #             ...

    #         case "Glosa - Recursar TST":
    #             ...

    #         case "Integralis - Enviar anexos Bradesco":
    #             ...

    #         case "Relatório - Brindes":
    #             ...

    #         case "Tesouraria - Nota Fiscal":
    #             ...
    #         case _:
    #             ...
            
        

    def bottonSair(self):
        self.botaoSair = customtkinter.CTkButton(self.tela,text="Sair", text_color="Red",fg_color="transparent",bg_color="transparent",width=80,command=lambda: threading.Thread(target=self.logOf()).start())
        self.botaoSair.pack(padx=0, pady=0)