import threading
import customtkinter
from PIL import Image, ImageTk
from Presentation.Desktop.Gif import ImageLabel
import Application.AppService.ChamarTabelaAppService as Tabela
import Application.AppService.ObterAutomacaoSetorAppService as AutomacoesAppService
from Presentation.Desktop.FaturasPreFaturadas import telaTabelaFaturas


class TelaSelecioneAutomacoes:
    def __init__(self, setorUsuario, janela,token):
        # self.bottonSair(janela)
        self.modoEscuroAut(janela)
        self.labelSelecioneAutomacao(janela)
        self.comboAutomacoes(setorUsuario, janela)
        self.botaoAvancar(janela,token)
        
    def logOf(self):
        # self.bottonSair.pack_forget()
        self.ocultarTelaInicial()
        self.reiniciarTelaLogin()

    def modoEscuroAut(self,janela):
        self.photo = customtkinter.CTkImage(light_image = Image.open(r"Infra\Arquivos\logo.png"), size=(80,90))
        self.botaoDark = customtkinter.CTkButton(janela,text="",image=self.photo, hover_color="White",fg_color="transparent",bg_color="transparent",command=lambda: threading.Thread(target=self.modoEscuro()).start())
        self.botaoDark.pack()

    def ocultarBotaoDark(self):
        self.botaoDark.pack_forget()

    def ocultarTelaInicial(self):
        self.iniciarAvancar.pack_forget()
        self.comboBoxAutomacao.pack_forget()
        self.labelAutomacao.pack_forget()

    def labelSelecioneAutomacao(self, janela):
        self.labelAutomacao = customtkinter.CTkLabel(janela, text="Selecione uma Automação:", text_color="#274360")
        self.labelAutomacao.pack(padx=0, pady=0)

    def comboAutomacoes(self,setorUsuario,janela):
        self.listaAutomacao = AutomacoesAppService.obterListaSetor(setorUsuario)
        self.comboBoxAutomacao = customtkinter.CTkComboBox(janela,values=self.listaAutomacao, border_color="#274360",button_color="#274360", button_hover_color="White",dropdown_fg_color="#274360", dropdown_text_color="White",text_color="#274360",width=300)
        self.comboBoxAutomacao.pack(padx=10, pady=25)

    def botaoAvancar(self, janela,token):
        self.iniciarAvancar = customtkinter.CTkButton(janela, width=80,fg_color="#274360",text="Avançar", command=lambda: threading.Thread(target=self.avancar(janela,token)).start())
        self.iniciarAvancar.pack(padx=10, pady=10)

    def avancar(self,janela,token):
        self.ocultarTelaInicial()
        # ImageLabel.iniciarGif(self,janela,texto="Aguarde...")
        automacaoSelecionada = self.comboBoxAutomacao.get()
        self.ocultarBotaoDark()

        match automacaoSelecionada:
            case "Faturamento - Enviar PDF BRB":
                convenio = 10
                telaTabelaFaturas(janela=janela,token=token,nomeAutomacao=automacaoSelecionada,codigoConvenio=convenio)
                # Tabela.aplicarChamada(self,janela,token,automacaoSelecionada)

    def bottonSair(self,janela):
        self.botaoSair = customtkinter.CTkButton(janela,text="Sair", text_color="Red",fg_color="transparent",bg_color="transparent",width=80,command=lambda: threading.Thread(target=self.logOf()).start())
        self.botaoSair.pack(padx=0, pady=0)