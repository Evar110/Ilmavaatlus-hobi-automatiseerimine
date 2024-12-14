from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service as FirefoxService
from bs4 import BeautifulSoup

andmed = []
url = "https://www.ilmateenistus.ee/ilm/ilmavaatlused/vaatlusandmed/tunniandmed/"

# Geckodriveri tee määramine
geckodriver_path = "C:/Program Files/geckodriver/geckodriver.exe"  # Asenda oma WebDriveri tee

service = FirefoxService(executable_path=geckodriver_path)
driver = webdriver.Firefox(service=service)

# Avab veebilehe
driver.get(url)

# Kuupäeva sisestamine
date_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "ood-datepicker"))) # ootab 20s või kuni saab lisada kuupäeva
date_input.clear()
date_input.send_keys('01.02.2024')
date_input.send_keys(Keys.RETURN)
        
# Kellaaja sisestamine
time_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "ood-timepicker"))) # ootab 20s või kuni saab lisada kellaaja
time_input.clear()
time_input.send_keys('7')
time_input.send_keys(Keys.RETURN)

# Ootab 20s või kuni leht lõpetab andmete laadimise
WebDriverWait(driver, 20).until(lambda d: driver.execute_script("return jQuery.active == 0"))

# Parsib HTML-i BeautifulSoup abil
html = driver.page_source
soup = BeautifulSoup(html, "html.parser")
tabel = soup.find("table", class_="table")

if tabel:
    read = tabel.find_all("tr")[1:]  # Jätab vahele päise rea
    for rida in read:
        andmete_rida = rida.find_all("td")
        if andmete_rida:
            asukoht = andmete_rida[0].get_text(strip=True)
            andmed.append(asukoht)

driver.quit()

with open('asukohad.txt', 'w',  encoding='utf-8') as sisu:
    sisu.write('[')
    sisu.write(', '.join(f'"{asukoht}"' for asukoht in andmed))
    sisu.write(']')