

def validarSetor(roules, setor):
    for role in roules:
        if (setor in role["Value"]):
            return True