import json
import time
import requests


def auth(login,senha):
    global token
    global proxies

    urlAuth = 'https://amhptiss.amhp.com.br/api/Auth'
    proxies = {
        "http": f"http:// + {login}:{senha}@10.0.0.230:3128/",
        "https": f"http://{login}:{senha}@10.0.0.230:3128/"
    }

    usuario_login = {
        "Usuario": login,
        "Senha" : senha
    }

    post = requests.post( urlAuth, usuario_login, proxies=proxies)
    content = json.loads(post.content)
    token = content['AccessToken']
    print('Token =>',token)

    # time.sleep(1)
    # print(post.text)
    # content = json.loads(post.content)
    # roles = content['UsuarioToken']["Claims"]
    # return roles