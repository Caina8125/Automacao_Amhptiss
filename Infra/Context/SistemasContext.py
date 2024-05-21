import time
from abc import ABC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class PageElement(ABC):
    def __init__(self, driver, url=''):
        self.driver = driver
        self.url = url
    def open(self):
        self.driver.get(self.url)

class Login(PageElement):
    def dadosLogin(self, xpathUsuario, xpathsenha,xpathEntrar,xpathOpcional):
        self.usuario = xpathUsuario
        self.senha = xpathsenha
        self.entrar = xpathEntrar
        self.opcional = xpathOpcional

    def logar(self, usuario, senha):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario).send_keys(usuario)
        time.sleep(2)
        self.driver.find_element(*self.senha).send_keys(senha)
        time.sleep(2)
        self.driver.find_element(*self.entrar).click()
        time.sleep(5)