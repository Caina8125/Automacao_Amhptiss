from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui
# Caminho para o executável do ChromeDriver
# chromedriver_path = r"C:\Automacao_Amhptiss\Infra\Arquivos"

# Configurações do navegador
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = r"C:\Program Files\MV\SoulMV\SoulMV.exe"  # Caminho para o executável do Chromium

# Inicializa o driver do Chrome
driver = webdriver.Chrome(options=chrome_options)

print(driver.window_handles)

url = "http://soulmv.gruposanta.com.br/mvautenticador-cas/login"
# time.sleep(5)
# pyautogui.press("TAB")
# time.sleep(1)
# pyautogui.typewrite(url)
# time.sleep(1)
# pyautogui.press("enter")

# time.sleep(5)

driver.switch_to.window(driver.window_handles[2])
print(driver.current_window_handle)

driver.get(url)
print(driver.window_handles)

time.sleep(2)
driver.find_element(By.ID, 'username').send_keys("lucas.timoteo")

time.sleep(6000)