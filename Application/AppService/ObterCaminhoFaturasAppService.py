import pandas as pd
import Infra.Repository.GuiasEscaneadasRepository as GuiasEscaneadas

def TrasformarDataFrame(lista):
    try:
        df = pd.DataFrame(lista)
        df.columns = ["Fatura", "Remessa", "Convenio","Usuario"]
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
                return GuiasEscaneadas.BuscarGuiasEscaneadas(faturasUsuario,caminho,"")
            case 433:
                caminho = r"\\10.0.0.239\guiasscaneadas\2024\GDF"
                variavelFatura = "LOTE "
                return GuiasEscaneadas.BuscarGuiasEscaneadas(faturasUsuario,caminho,"")
        

def IniciarBusca(lista,codigoConvenio):
    df = TrasformarDataFrame(lista)
    # robo = ObterCaminhosFaturaRobo(df)ks
    usuario = ObterCaminhosFaturaUsuario(df,codigoConvenio)
    return usuario

  