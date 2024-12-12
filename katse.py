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

# Kasutaja sisend
kuupäev = input("Sisesta kuupäev (formaadis pp.kk.aaaa): ")
hommikune_aeg = input("Sisesta hommikune kellaaeg (ainult täistund, formaadis tt:00): ")
ohtune_aeg = input("Sisesta õhtune kellaaeg (ainult täistund, formaadis tt:00): ")
print('\n')

# Käivitab Firefox WebDriver
service = FirefoxService(executable_path=geckodriver_path)
driver = webdriver.Firefox(service=service)

# funktsioon andmete kogumiseks
def andmete_kogumine(kellaaeg, kogutav):
    """
    :parameeter kellaaeg: Kellaaeg, mida kasutada päringus (näiteks "7:00").
    :parameeter kogutav: Andmed, mida koguda ('hommik' või 'õhtu').
    """
    try:
        # Avab veebilehe
        driver.get(url)

        # Kuupäeva sisestamine
        date_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "ood-datepicker"))
        )
        date_input.clear()
        date_input.send_keys(kuupäev)
        date_input.send_keys(Keys.RETURN)
        
        # Kellaaja sisestamine
        time_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "ood-timepicker"))
        )
        time_input.clear()
        time_input.send_keys(kellaaeg)
        time_input.send_keys(Keys.RETURN)

        # Ootab, kuni leht lõpetab andmete laadimise
        WebDriverWait(driver, 20).until(
            lambda d: driver.execute_script("return jQuery.active == 0")
        )

        # Parsib HTML-i BeautifulSoup abil
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        tabel = soup.find("table", class_="table")

        if tabel:
            read = tabel.find_all("tr")[1:]  # Jätab vahele päise rida
            for rida in read:
                andmete_rida = rida.find_all("td")
                if andmete_rida:
                    asukoht = andmete_rida[0].get_text(strip=True)
                    temperatuur = andmete_rida[1].get_text(strip=True)
                    õhurõhk = andmete_rida[3].get_text(strip=True)
                    tuule_suund = andmete_rida[5].get_text(strip=True)

                    if asukoht == "Võru":  # Filtreerib valitud asukoha andmed
                        print(f"Asukoht: {asukoht}")
                        print(f"Temperatuur: {temperatuur if temperatuur else '-'}")
                        andmed.append(temperatuur if temperatuur else '-')

                        if kogutav == "õhtu":
                            print(f"Õhurõhk: {õhurõhk if õhurõhk else '-'}")
                            print(f"Tuule suund: {tuule_suund if tuule_suund else '-'}")
                            andmed.extend([
                                õhurõhk if õhurõhk else '-',
                                tuule_suund if tuule_suund else '-'
                            ])
                        print("-" * 40)
    except Exception as e:
        print(f"Tekkis viga: {e}")

# Andmete kogumine hommikul (ainult temperatuur)
andmete_kogumine(hommikune_aeg, "hommik")

# Andmete kogumine õhtul (temperatuur, õhurõhk, tuule suund)
andmete_kogumine(ohtune_aeg, "õhtu")

# Lõpeta WebDriver
driver.quit()

print("Kogutud andmed:", andmed)