import Application.AppService.EnviarPdfBrb as AnexarBrb

def aplicarChamada(automacao):
    match automacao:
        case "Faturamento - Enviar PDF BRB":
            AnexarBrb.enviar_pdf()
