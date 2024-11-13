   # Asjad mida võibolla lisada/muuta/parandada:
            ## 1. Lasta kasutajal valida:
                ###1.1. Hommikuse salvestamise limiidi aja
                ###1.2. Faili salvestamise koha + nime

                #Probleem on selles, et sellega läheks kood 2-5x suuremaks kasutaja errorite pärast
                #Ei saa kasutaja erroreid lihtsalt ignoda ka, kuna muid exceli fail corruptib ilma backup võimaluseta
                    #Mul juhtunud vähemalt 5x

            ##2. Leida parem koht kust infot saada.
                #Praegune tundub vääääääääga halb, aga see oli ainukene mille leidsin kus ei olnud tableid hidden
                #Tavalisel ilmateenistus.ee lehel ei ole JSON faile mida arvuto saaks kergelt lugeda ja ei leidnud kust nad info saavad.
                    ##Mingi lehe leidsin, aga see oli veel keerulisem.
                ##Ma ei tahtnud mingit eraldi äppi alla laadida, mis seda teeks, kuna see teeks kasutaja elu raskemaks

            ##3. Programm corruptib exceli faili kui seal kasutaja liiga palju ise muudab.
                #Tõenäoliselt kui lisada samal viisil nagu käsitsi uus kuu või uus aasta siis peale seda corruptiks prg

            ##4. Vaja lisada nupp 'backup' mis teeks koopia algsest exceli failist. Vb seda automatiseerida, sest corruptib palju.
                ###Nt, enne igat 'write' käsku teeb koopia ja kui edukat 'write' läbitud kustuab vana koopia kui seda on

            ##5. Parandada uue kuu ja uue aasta kirjutamine
                #Peaks olema üsna kerge, ma lihtsalt ei tea kuidas koodiga excelis nii kirjutada

            ##6. FutureWarning: The behavior of DataFrame concatenation with empty or all-NA entries is deprecated.
                 #In a future version, this will no longer exclude empty or all-NA columns when determining the result dtypes.
                 #To retain the old behavior, exclude the relevant entries before the concat operation.
                 #lõplik_df = pd.concat([olemasolev_df, uued_andmed_df], ignore_index=True)
                    ##Suva praegu

            ##7. Juhul, kui hommikused andmed jäid panemata, tegin nii, et ei saa õhtuseid ka panna v.a kui faili lood.
                #Band-aid fix, peaks tegelikult leidma viisi kuidas hommikusi andmeid ikka leida või midagi sarnast
                #Peamine asi on kuna ma hard-codisin sisse, et õhtul saab ainult õhu temperatuuri lisada ja ei viitsinud muuta

            ##8. Kui valida juba varem tehtud fail kus on midagi kirjas siis see corruptib selle
