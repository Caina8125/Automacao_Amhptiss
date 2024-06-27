from Presentation.Desktop.FaturasPreFaturadas import telaTabelaFaturas
from Presentation.Desktop.SelecaoUmArquivo import SelecaoUmArquivo
from Presentation.Desktop.SelecioneAutomacao import TelaSelecioneAutomacoes
import Application.AppService.AutomacaoService as automacoes
import tkinter as tk
import customtkinter

from Presentation.Desktop.TelaBuscaProtocolo import tela_busca_protocolo

class SelecionarFaturasPre(TelaSelecioneAutomacoes, telaTabelaFaturas, tela_busca_protocolo):
    def __init__(self, setorUsuario, janela, token):
        self.iniciar_tela_selecione(setorUsuario, janela, token)

    def avancar(self,token):
        # ImageLabel.reiniciarGif(self,janela=self.tela,texto="Buscando faturas no \nAMHPTISS...")
        # self.ocultaBotaoAvancar()
        # time.sleep(5)
        self.ocultarTelaInicial()
        automacaoSelecionada = self.comboBoxAutomacao.get()
        self.ocultarBotaoDark()

        match automacaoSelecionada:
            case "Faturamento - Anexar Guia Geap":
                self.tela_tabela_faturas(janela=self.tela,token=token,obj=automacoes.Geap('https://www2.geap.com.br/auth/prestadorVue.asp', '66661692120', 'Amhp2024', 'normal'),codigoConvenio=225, setor=self.setorUsuario)

            case "Faturamento - Conferência de Protocolos":
                ...

            case "Faturamento - Conferência GEAP":
                ...

            case "Faturamento - Conferência Bacen":
                ...

            case "Faturamento - Conferência Envio GDF":
                self.tela_busca_protocolo(janela=self.tela,token=token,obj=automacoes.Geap('https://www2.geap.com.br/auth/prestadorVue.asp', '66661692120', 'Amhp2024', 'normal'),codigoConvenio=433, setor=self.setorUsuario)

            case "Faturamento - Enviar PDF Bacen":
                ...

            case "Faturamento - Enviar PDF Benner":
                ...

            case "Faturamento - Enviar PDF BRB":
                ...
                # telaTabelaFaturas(janela=self.tela,token=token,obj=automacoes.EnviarPdf(),codigoConvenio=10, setor=self.setorUsuario)

            case "Faturamento - Enviar PDF GDF":
                self.tela_tabela_faturas(janela=self.tela,token=token,obj=automacoes.NextCloudMaidaEnvio('https://nextcloud.maida.health/login', 433, '735860000173', 'nopaperpass'),codigoConvenio=433, setor=self.setorUsuario)
            
            case "Faturamento - Enviar PDF PMDF":
                self.tela_tabela_faturas(janela=self.tela,token=token,obj=automacoes.NextCloudMaidaEnvio('https://nextcloud.maida.health/login', 381, '735860000173', 'nopaperpass'),codigoConvenio=381, setor=self.setorUsuario)

            case "Faturamento - Enviar PDF SIS":
                self.tela_tabela_faturas(janela=self.tela,token=token,obj=automacoes.NextCloudMaidaEnvio('https://nextcloud.maida.health/login', 160, '735860000173', 'nopaperpass'),codigoConvenio=160, setor=self.setorUsuario)

            case "Faturamento - Enviar XML Bacen":
                ...

            case "Faturamento - Enviar XML Benner":
                ...

            case "Faturamento - Enviar XML Caixa":
                ...

            case "Faturamento - Leitor de PDF GAMA":
                ...

            case "Faturamento - Verificar Situação BRB":
                ...

            case "Faturamento - Verificar Situação Fascal":
                ...

            case "Faturamento - Verificar Situação Gama":
                ...

            case "Financeiro - Buscar Faturas GEAP":
                ...

            case "Financeiro - Demonstrativos Amil":
                ...

            case "Financeiro - Demonstrativos BRB":
                ...

            case "Financeiro - Demonstrativos Câmara dos Deputados":
                ...

            case "Financeiro - Demonstrativos Camed":
                ...

            case "Financeiro - Demonstrativos Casembrapa":
                ...

            case "Financeiro - Demonstrativos Cassi":
                ...

            case "Financeiro - Demonstrativos Codevasf":
                ...

            case "Financeiro - Demonstrativos E-Vida":
                ...

            case "Financeiro - Demonstrativos Fapes":
                ...

            case "Financeiro - Demonstrativos Fascal":
                ...

            case "Financeiro - Demonstrativos Gama":
                ...

            case "Financeiro - Demonstrativos Life Empresarial":
                ...

            case "Financeiro - Demonstrativos MPU":
                ...

            case "Financeiro - Demonstrativos PMDF":
                ...

            case "Financeiro - Demonstrativos Postal":
                ...

            case "Financeiro - Demonstrativos Real Grandeza":
                ...

            case "Financeiro - Demonstrativos Saúde Caixa":
                ...

            case "Financeiro - Demonstrativos Serpro":
                ...

            case "Financeiro - Demonstrativos SIS":
                ...

            case "Financeiro - Demonstrativos STF":
                ...

            case "Financeiro - Demonstrativos TJDFT":
                ...

            case "Financeiro - Demonstrativos Unafisco":
                ...

            case "Glosa - Gerador de Planilha GDF":
                ...

            case "Glosa - Gerar Planilhas SERPRO":
                ...

            case "Glosa - Filtro Matrículas":
                ...

            case "Glosa - Recursar Amil":
                ...  

            case "Glosa - Recursar Benner(Câmara, CAMED, FAPES, Postal)":
                ...

            case "Glosa - Recursar BRB":
                ...
                # SelecaoUmaPasta(self.tela, AppService.RecursoBrb('https://portal.saudebrb.com.br/GuiasTISS/Logon', '00735860000173', 'AMHP7356!'))

            case "Glosa - Recursar Casembrapa":
                ...

            case "Glosa - Recursar Cassi":
                ...

            case "Glosa - Recursar GEAP Duplicado":
                ...

            case "Glosa - Recursar GEAP Sem Duplicado":
                ...
            
            case "Glosa - Recursar E-VIDA":
                ...

            case "Glosa - Recursar Fascal":
                ...

            case "Glosa - Recursar Gama":
                ...

            case "Glosa - Recursar Petrobras":
                ...

            case "Glosa - Recursar Real Grandeza":
                ...

            case "Glosa - Recursar Saúde Caixa":
                ...

            case "Glosa - Recursar SIS":
                ...

            case "Glosa - Recursar STF":
                ...

            case "Glosa - Recursar STM":
                ...

            case "Glosa - Recursar TJDFT":
                ...

            case "Glosa - Recursar TST":
                ...

            case "Integralis - Enviar anexos Bradesco":
                ...

            case "Relatório - Brindes":
                ...

            case "Tesouraria - Nota Fiscal":
                SelecaoUmArquivo(self.tela, )
            case _:
                ...

    def botao_voltar_click(self):
        self.container1.pack_forget()
        self.container2.pack_forget()
        self.container3.pack_forget()
        self.container4.pack_forget()
        self.container5.pack_forget()
        self.botaoVoltar.place_forget()
        try:
            self.botaoSelectAll.place_forget()
        except:
            pass
        self.iniciar_tela_selecione(self.setorUsuario, self.tela, self.token)

    def ocultar_botoes_enviar_remover(self):
        try:
            self.botaoStart.place_forget()
            self.botaoStart.place_forget()
            self.botaoRemover.destroy()
            self.botaoRemover.destroy()
            self.botaoStart = customtkinter.CTkButton(tk.Frame())
            self.botaoStart.pack()
            self.botaoStart.pack_forget()
            self.botaoRemover = customtkinter.CTkButton(tk.Frame())
            self.botaoRemover.pack()
            self.botaoRemover.pack_forget()

            self.botaoRemover = None
        except:
            pass
        