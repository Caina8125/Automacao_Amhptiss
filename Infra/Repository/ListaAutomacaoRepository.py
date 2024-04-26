def obterListaFaturamento():
    automacaoFaturamento = ["Faturamento - Anexar Honorario Geap",
                            "Faturamento - Conferência GEAP",
                            "Faturamento - Conferência Bacen",
                            "Faturamento - Enviar PDF Bacen",
                            "Faturamento - Enviar PDF BRB",
                            "Faturamento - Enviar XML Bacen",
                            "Faturamento - Verificar Situação BRB",
                            "Faturamento - Verificar Situação Fascal",
                            "Faturamento - Verificar Situação Gama",]
    return automacaoFaturamento

def obterListaFinanceiro():
    automacaoFinanceiro = [ "Financeiro - Buscar Faturas GEAP", 
                            "Financeiro - Demonstrativos BRB", 
                            "Financeiro - Demonstrativos Câmara dos Deputados", 
                            "Financeiro - Demonstrativos Camed", 
                            "Financeiro - Demonstrativos Casembrapa", 
                            "Financeiro - Demonstrativos Cassi", 
                            "Financeiro - Demonstrativos Codevasf", 
                            "Financeiro - Demonstrativos E-Vida", 
                            "Financeiro - Demonstrativos Fapes", 
                            "Financeiro - Demonstrativos Fascal", 
                            "Financeiro - Demonstrativos Gama", 
                            "Financeiro - Demonstrativos MPU", 
                            "Financeiro - Demonstrativos PMDF", 
                            "Financeiro - Demonstrativos Postal", 
                            "Financeiro - Demonstrativos Saúde Caixa", 
                            "Financeiro - Demonstrativos Serpro", 
                            "Financeiro - Demonstrativos SIS", 
                            "Financeiro - Demonstrativos STF", 
                            "Financeiro - Demonstrativos TJDFT", 
                            "Financeiro - Demonstrativos Unafisco", ]
    return automacaoFinanceiro

def obterListaGlosa():
    automacaoGlosa = ["Glosa - Gerador de Planilha GDF",
                            "Glosa - Recursar GEAP Duplicado",
                            "Glosa - Recursar GEAP Sem Duplicado",
                            "Glosa - Recursar Postal",
                            "Glosa - Recursar Saúde Caixa",
                            "Glosa - Recursar Serpro",
                            "Glosa - Recursar SIS",]
    return automacaoGlosa

def obterListaTesouraria():
    automacaoTesouraria = ["Tesouraria - Nota Fiscal"]
    return automacaoTesouraria