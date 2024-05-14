
def validarSetor(roules, setor):
    for role in roules:
        if (setor in role["Value"] or role["Value"] == "Administrador"):
            return True