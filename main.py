import pandas as pd
import os

print("Traitement des données en cours ...")
if os.path.exists("exemple.xlsx"):  # vérification que le fichier existe dans le répertoire
    # Selection des colonnes qu'on souhaite traiter
    df = pd.read_excel("exemple.xlsx", usecols="A,C:E,F")  # A:Etapes, C:Nom, I à K: NDI + tranche SDA, N: GA

    # Filtrage par nom d'étape et ne pas pendre le numéro du GA absent
    df = df[df["Etapes"] == "En attente routage SBC"]
    df = df[df["GA"].notnull()]

    # Récupération de toutes les données de chaques colonnes
    ndi_frais = df["NDI FRAIS"].values
    Ga = df["GA"].values
    start_sda = df["Début SDA"].values
    end_sda = df["Fin SDA"].values

    # Déclaration de nouvelles listes pour le traitement des données
    liste_sda = []
    liste_ndi = []
    liste_ga = []

    # récupération et traitement de chaques champs de chaques colones
    for s_sda, e_sda, ndi, ga in zip(start_sda, end_sda, ndi_frais, Ga):
        end_int = s_sda % 10  # récuperer le dernier chiffre de la SDA
        if end_int == 1:
            s_sda -= 1
        plage_sda = (e_sda - s_sda + 1)
        nb_cent = plage_sda / 100

        # boucle pour traiter si le nombre de SDA est supérieur ou égale à 100
        if nb_cent >= 1:
            if end_int in [0, 1]:  # vérification du dernier chiffre de la SDA
                nb_cent = int(plage_sda / 100)
                s_sda_cent = str(s_sda)
                s_sda_cent = int(s_sda_cent[:-2])

                # Traitement du (ou des) centaine(s) de SDA
                for _ in range(nb_cent):
                    liste_sda.append(s_sda_cent)
                    liste_ndi.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI={ndi}!")
                    liste_ga.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")
                    s_sda_cent += 1

                # Traitement des dixaines 
                rest_cent = plage_sda % 100
                if rest_cent >= 10:
                    rest_dix = int(rest_cent % 10)
                    nb_dix = int(rest_cent / 10)
                    s_sda_dix = s_sda_cent * 10
                    for _ in range(nb_dix):
                        liste_sda.append(s_sda_dix)
                        liste_ndi.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI={ndi}!")
                        liste_ga.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")
                        s_sda_dix += 1

                    # Traitement des unités
                    if 0 < rest_dix < 10:
                        s_sda_dix *= 10
                        for _ in range(rest_dix + 1):
                            liste_sda.append(s_sda_dix)
                            liste_ndi.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI={ndi}!")
                            liste_ga.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")
                            s_sda_dix += 1

            else:
                # Traitement de(s) unité(s)
                nb_unit = plage_sda % 10
                for _ in range(nb_unit):
                    liste_sda.append(s_sda)
                    liste_ndi.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI={ndi}!")
                    liste_ga.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")
                    s_sda += 1

                # Traitement de(s) dixaine(s)
                nb_dix = int(((plage_sda - nb_unit) / 10) % 10)
                if nb_dix >= 1:
                    s_sda_dix = s_sda // 10
                    for _ in range(nb_dix):
                        liste_sda.append(s_sda_dix)
                        liste_ndi.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI={ndi}!")
                        liste_ga.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")
                        s_sda_dix += 1

                    # Traitement de(s) centaine(s)
                    nb_cent = plage_sda // 100
                    if nb_cent >= 1:
                        s_sda_cent = s_sda_dix // 10
                        for _ in range(nb_cent):
                            liste_sda.append(s_sda_cent)
                            liste_ndi.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI={ndi}!")
                            liste_ga.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")
                            s_sda_cent += 1

        else:
            # Traitement si le nombre de SDA est inférieur à 100
            nb_dix = plage_sda // 10
            rest_dix = int(plage_sda % 10)
            if plage_sda >= 10:
                s_sda_dix = str(s_sda)
                s_sda_dix = int(s_sda_dix[:-1])
                for _ in range(nb_dix):
                    liste_sda.append(s_sda_dix)
                    liste_ndi.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI={ndi}!")
                    liste_ga.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")
                    s_sda_dix += 1
                if rest_dix < 10:
                    s_sda_dix *= 10
                    for _ in range(rest_dix):
                        liste_sda.append(s_sda_dix)
                        liste_ndi.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI={ndi}!")
                        liste_ga.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")
                        s_sda_dix += 1

            else:
                for _ in range(plage_sda):
                    liste_sda.append(s_sda)
                    liste_ndi.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI={ndi}!")
                    liste_ga.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")
                    s_sda += 1

    # Convertion de la liste de Sda en chaine de caractère + ajouter le 0 au début
    Liste_Sda = [f"0{str(sda)}" for sda in liste_sda]

    # création d'un dossier sauf si il existe déjà
    os.makedirs('imports/', exist_ok=True)
    print("Création des fichers d'imports de masse dans le dossier imports: En cours ...")

    # Création des fichiers d'import de masse
    df_10D = pd.DataFrame({"OPERATION": "Add", "PUBID": Liste_Sda, "SED": liste_ga})
    df_10D.to_csv('imports/import_routage_SDA.csv', index=False)

    df_ndi = pd.DataFrame({"OPERATION": "Add", "PUBID": Liste_Sda, "SED": liste_ndi})
    df_ndi.to_csv('imports/import_NDI_Frais.csv', index=False)

    print("Création des fichers d'imports de masse dans le dossier imports: Terminé")

else:
    print("Le fichier n'existe pas ou le nom du fichier n'est pas correct.")
    print("Nom du fichier: FICHIER DE SUIVI T-SIP_v2.xlsx")
    print("Merci de vérifier le nom du fichier ou la présence du fichier dans le même répertoire que l'exécutable.")
os.system("pause")
