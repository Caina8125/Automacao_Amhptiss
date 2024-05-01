import threading
import customtkinter
import Presentation.Desktop.Gif as gif
import Application.AppService.ChamarAutomacoesAppService as chamarAutomacoes
import Application.AppService.ObterAutomacaoSetorAppService as AutomacoesAppService
import Presentation.Desktop.FaturasPreFaturadas as TelaFaturas

class TelaSelecioneAutomacoes:
    def __init__(self, setorUsuario, janela):
        # self.bottonSair(janela)
        self.comboAutomacoes(setorUsuario, janela)
        self.botaoAvancar(janela)
        
    def logOf(self):
        # self.bottonSair.pack_forget()
        self.ocultarTelaInicial()
        self.reiniciarTelaLogin()

    def ocultarTelaInicial(self):
        self.iniciarAutomacao.pack_forget()
        self.comboBoxAutomacao.pack_forget()

    def comboAutomacoes(self,setorUsuario,janela):
        self.listaAutomacao = AutomacoesAppService.obterListaSetor(setorUsuario)
        self.comboBoxAutomacao = customtkinter.CTkComboBox(janela,values=self.listaAutomacao, border_color="#274360",button_color="#274360", button_hover_color="White",dropdown_fg_color="#274360", dropdown_text_color="White",text_color="#274360",width=300)
        self.comboBoxAutomacao.pack(padx=10, pady=25)

    def botaoAvancar(self, janela):
        self.iniciarAvancar = customtkinter.CTkButton(janela, width=80,fg_color="#274360",text="Avançar", command=lambda: threading.Thread(target=self.chamarAutomacao()).start())
        self.iniciarAvancar.pack(padx=10, pady=10)

    def avancar(self):
        self.ocultarTelaInicial()
        self.gif("Trabalhando...")
        automacaoSelecionada = self.comboBoxAutomacao.get()
        chamarAutomacoes.aplicarChamada(automacaoSelecionada)

    def bottonSair(self,janela):
        self.botaoSair = customtkinter.CTkButton(janela,text="Sair", text_color="Red",fg_color="transparent",bg_color="transparent",width=80,command=lambda: threading.Thread(target=self.logOf()).start())
        self.botaoSair.pack(padx=0, pady=0)