 
# mod to wait load page
import time

# webdriver
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

# select mod
from selenium.webdriver.support.ui import Select

# 
from bs4 import BeautifulSoup

#Files
import shutil

# SqlServer
#import pyodbc

options = Options()
options.add_experimental_option("prefs", {
  "download.default_directory": r"\\172.30.5.2\c$\Projects\Python\Files",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})
options.add_argument("--start-maximized")

driver = webdriver.Chrome(chrome_options=options)

#Path to download
try:
    shutil.rmtree(r'\\172.30.5.2\c$\Projects\Python\Files')
except: 
  pass

#to site to access
driver.get("https:\\www.yoursite...")

# sleep to load page
time.sleep(3)

# get the element
user = driver.find_element_by_id("UserName")
senha = driver.find_element_by_id("Senha")
login = driver.find_element_by_id("botaoLogin")

#USer and password
user.send_keys("user")
senha.send_keys("password")

# click in login
login.click()

# o Sleep abaixo e para aguardar o carregamento da pagina
time.sleep(3)

#Right
#setaDireta = driver.find_element_by_class_name("perfil")
#Ex: <p class="p0 ng-binding">text</p>

setaDireta = driver.find_element_by_css_selector('a.arrow-right.widgetScrollerRight')

setaDireta.click()

#down
botoes = driver.find_elements_by_css_selector('div.widgets-vermais')

c = 0
for botao in botoes:
    if c == 3:
        botao.click()
    c = c + 1

# o Sleep abaixo e para aguardar o carregamento da pagina
time.sleep(10)

#Focus in the button and make scroll down
element = driver.find_element_by_id("botaoExportarSolicitacoesXls")
actions = ActionChains(driver)
actions.move_to_element(element).perform()

#Click export button
exportar = driver.find_element_by_id("botaoExportarSolicitacoesXls")
exportar.click()

# sleep to wait export csv
time.sleep(10)

#close drive
driver.close()


