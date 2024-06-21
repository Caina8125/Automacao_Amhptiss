import pandas as pd
import Infra.Repository.GuiasEscaneadasRepository as GuiasEscaneadas

def TrasformarDataFrame(lista):
    try:
        df = pd.DataFrame(lista)
        df.columns = ["Fatura", "Remessa", "Convenio","Usuario", "Protocolo", "Envio Operadora"]
        return df
    except Exception as e:
        print(e)

def ObterCaminhosFaturaRobo(dataFrame):
    faturasRobo = dataFrame.loc[(dataFrame["Usuario"] == "root")]
    if not faturasRobo.empty:
        pass

def ObterCaminhosFaturaUsuario(dataFrame,codigoConvenio):
    faturasUsuario = dataFrame.loc[(dataFrame["Usuario"] != "root")]
    if not faturasUsuario.empty:
        match codigoConvenio:
            case 10:
                caminho = r"\\10.0.0.239\guiasscaneadas\2024\BRB"
                return GuiasEscaneadas.BuscarGuiasEscaneadas(faturasUsuario,caminho, codigoConvenio)
            case 433:
                caminho = r"\\10.0.0.239\guiasscaneadas\2024\GDF"
                return GuiasEscaneadas.BuscarGuiasEscaneadas(faturasUsuario,caminho, codigoConvenio)
            case 225:
                caminho = r"\\10.0.0.239\guiasscaneadas\2024\GEAP"
                return GuiasEscaneadas.BuscarGuiasEscaneadas(faturasUsuario,caminho, codigoConvenio)
        

def IniciarBusca(lista,codigoConvenio):
    df = TrasformarDataFrame(lista)
    # robo = ObterCaminhosFaturaRobo(df)ks
    usuario = ObterCaminhosFaturaUsuario(df,codigoConvenio)
    return usuario

  