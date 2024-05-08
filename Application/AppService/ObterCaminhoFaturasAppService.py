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

def ObterCaminhosFaturaUsuario(dataFrame):
    faturasUsuario = dataFrame.loc[(dataFrame["Usuario"] != "root")]
    if not faturasUsuario.empty:
        return GuiasEscaneadas.GuiasEscaneadasBrb(faturasUsuario)
        

def IniciarBusca(lista):
    df = TrasformarDataFrame(lista)
    # robo = ObterCaminhosFaturaRobo(df)
    usuario = ObterCaminhosFaturaUsuario(df)
    return usuario

  