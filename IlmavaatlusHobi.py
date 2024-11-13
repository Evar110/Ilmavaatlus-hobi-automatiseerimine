import tkinter as tk
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import os



def saada_ilma_andmed(url):
    vastus = requests.get(url)
    soup = BeautifulSoup(vastus.text, 'html.parser')
    ilma_andmed = []
    tabel = soup.find('table')
    if tabel:
        read = tabel.find_all('tr')
        if len(read) > 2:
            andme_rida = read[2].find('td')
            if andme_rida:
                tekst = andme_rida.get_text(separator="\n").strip()
                read = tekst.split('\n')
                temperatuur = read[0].split(':')[1].strip().replace('°C', '')
                tuul = read[1].split(':')[1].strip()
                tuule_suund = tuul.split(' ')[0]
                rõhk = read[3].split(':')[1].strip().replace('hPa', '')
                ilma_andmed.append([temperatuur, rõhk, tuule_suund])
            else:
                print("Ilma andmete rida ei leitud. Vaata üle HTMLi struktuur.")
        else:
            print("Ilma andmete tabelit ei leitud. Vaata üle HTMLi struktuur.")
    return ilma_andmed



def paskalites_mmHg(paskalid):
    return float(paskalid) * 0.00750062



def päev():
    eesti_päevad = {
        'Monday': 'E',
        'Tuesday': 'T',
        'Wednesday': 'K',
        'Wednesday': 'K',
        'Thursday': 'N',
        'Friday': 'R',
        'Saturday': 'L',
        'Sunday': 'P'
    }
    täna = datetime.today()
    nädalapäev = eesti_päevad[täna.strftime('%A')]
    päev_info = täna.strftime(f'%d.{nädalapäev}')
    return päev_info



def on_hommik(): #W.I.P
    praegune_tund = datetime.now().hour
    return praegune_tund < 12  #Kui seda muuta, siis muutub aeg millal kood laseb hommikuse info kirjutada, peale seda muutub õhtuseks



def põhifunktsioon():
    url = 'https://www.ilmateenistus.ee/ilma_andmed/ticker/vaatlused-html.php?jaam=15&stiil=1'
    #See vaatab ainult Tartu-Tõravere infot. Muuta saab kui muuta jaam=<number>, saaks automatiseerida
    ilma_andmed = saada_ilma_andmed(url)
    if not ilma_andmed:
        print("Pole andmeid mida salvestada. Lõpetamine.")
        return

    faili_tee = r'C:\Games\ilmavaatlus.xlsx'  #W.I.P, Faili salvestamise tee + nimi
    päev_info = päev()
    töödeldud_andmed = []

    # W.I.P, ei tööta praegu
    täna = datetime.today()
    if täna.day == 1:
        uus_periood = f"Uus kuu: {täna.strftime('%B')}"
    elif täna.month == 1 and täna.day == 1:
        uus_periood = f"Uus aasta: {täna.year}"
    else:
        uus_periood = None


    if on_hommik():
        temperatuur, rõhk, tuule_suund = ilma_andmed[0]
        rõhk_mmHg = round(paskalites_mmHg(rõhk), 2)
        töödeldud_andmed.append([päev_info, float(temperatuur), None, rõhk_mmHg, tuule_suund, uus_periood]) #None on õhtu temp koht
    else:
        temperatuur = ilma_andmed[0][0]
        töödeldud_andmed.append([päev_info, None, float(temperatuur), None, None, None]) #Õhtu

    #Exceli faili n.ö pealkirjade nimed (df-datafail)
    uued_andmed_df = pd.DataFrame(töödeldud_andmed, columns=[
        'Kuupäev', 'Hommik Temperatuur', 'Õhtu Temperatuur', 'Õhurõhk (mmHg)', 'Tuule suund', 'Uue perioodi märkus'])

    if os.path.exists(faili_tee):
        olemasolev_df = pd.read_excel(faili_tee)
        olemasolev_df['Hommik Temperatuur'] = olemasolev_df['Hommik Temperatuur'].astype(float)
        olemasolev_df['Õhtu Temperatuur'] = olemasolev_df['Õhtu Temperatuur'].astype(float)

        if on_hommik(): #Toimib kuid vaja teha arusaadavamaks
            uued_andmed_df = uued_andmed_df.dropna(how='all')
            if not ((olemasolev_df['Kuupäev'] == päev_info) & olemasolev_df['Hommik Temperatuur'].notnull()).any():
                lõplik_df = pd.concat([olemasolev_df, uued_andmed_df], ignore_index=True)
            else:
                print("Selle päeva hommikune info on juba lisatud. Kood ei lisanud midagi.")
                return
        else:
            mask = (olemasolev_df['Kuupäev'] == päev_info) & olemasolev_df['Õhtu Temperatuur'].isnull()
            if mask.any():
                olemasolev_df.loc[mask, 'Õhtu Temperatuur'] = uued_andmed_df.loc[0, 'Õhtu Temperatuur'] #.loc -pandas
                lõplik_df = olemasolev_df
            else:
                if ((olemasolev_df['Kuupäev'] == päev_info) & olemasolev_df['Õhtu Temperatuur'].notnull()).any():
                    print("Selle päeva õhtune info on juba lisatud. Kood ei lisanud midagi.")
                else:
                    print("Ei leidnud selle päeva hommikusi andmeid. Lõpetamine.")
                return
    else:
        lõplik_df = uued_andmed_df

    with pd.ExcelWriter(faili_tee, engine='openpyxl') as kirjutaja:
        lõplik_df.to_excel(kirjutaja, index=False)
    print(f'Fail edukalt salvestatud: {faili_tee}!')



def käivita_kood():
    põhifunktsioon()
    messagebox.showinfo("Kirjuta Exceli", "Programm toimis!")

#UI nupp
root = tk.Tk()
root.title("Ilmavaatlus")
execute_button = tk.Button(root, text="Kirjuta Exceli", command=käivita_kood)
execute_button.pack(pady=20)
root.mainloop()