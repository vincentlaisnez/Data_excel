import pandas as pd
import os

# Déclaration de listes pour le traitement des données
LISTE_SDA = []
LISTE_NDI = []
LISTE_GA = []


def update_list(sda, ndi, ga):
    """
    Mise à jour des listes LISTE_SDA, LISTE_NDI, et LISTE_GA avec les valeurs données.

    Args:
        sda (int): Sda ajouter à LISTE_SDA.
        ndi (str): ndi ajouter à LISTE_NDI avec le regex adapté.
        ga (str): Numéro de GA ajouter à LISTE_NDI avec le regex adapté.

    Returns:
        None
    """

    LISTE_SDA.append(f"0{sda}")
    LISTE_NDI.append(fr"!(^.*$)!sip:\1@SAG-OPER-URGENCE;PAI=0{ndi}!")  # f = f-string , r = Raw f-strings
    LISTE_GA.append(fr"!(^.*$)!sip:\1@SA-Aastra-ToIP{ga}-10!")


print("Traitement des données en cours ...")
if os.path.exists("exemple.xlsx"):  # vérification que le fichier existe dans le répertoire
    # Selection des colonnes qu'on souhaite traiter
    df = pd.read_excel("exemple.xlsx", usecols="A,C:E,F")  # A:Etapes, C à E: NDI + tranche SDA, F: GA

    # Filtrage par nom d'étape et ne pas pendre le numéro du GA absent
    df = df[df["Etapes"] == "En attente routage SBC"]
    df = df[df["GA"].notnull()]

    # Récupération de toutes les données de chaques colonnes
    ndi_frais = df["NDI FRAIS"].values
    ga = df["GA"].values
    start_sda = df["Début SDA"].values
    end_sda = df["Fin SDA"].values

    # récupération et traitement de chaques champs de chaques colones
    for s_sda, e_sda, ndi, ga in zip(start_sda, end_sda, ndi_frais, ga):
        end_int = s_sda % 10  # récuperer le dernier chiffre de la SDA
        if end_int == 1:
            s_sda -= 1
        plage_sda = (e_sda - s_sda + 1)
        nb_cent = plage_sda / 100

        # boucle pour traiter si le nombre de SDA est supérieur ou égale à 100
        if nb_cent >= 1:
            if s_sda % 100 == 0:  # vérification du dernier chiffre de la SDA
                nb_cent = int(plage_sda / 100)
                s_sda_cent = str(s_sda)
                s_sda_cent = int(s_sda_cent[:-2])

                # Traitement du (ou des) centaine(s) de SDA
                for _ in range(nb_cent):
                    update_list(s_sda_cent, ndi, ga)
                    s_sda_cent += 1

                # Traitement des dixaines 
                rest_cent = plage_sda % 100
                if rest_cent >= 10:
                    rest_dix = int(rest_cent % 10)
                    nb_dix = int(rest_cent / 10)
                    s_sda_dix = s_sda_cent * 10
                    for _ in range(nb_dix):
                        update_list(s_sda_dix, ndi, ga)
                        s_sda_dix += 1

                    # Traitement des unités
                    if 0 < rest_dix < 10:
                        s_sda_dix *= 10
                        sda_unit = s_sda_dix
                        for _ in range(rest_dix + 1):
                            update_list(sda_unit, ndi, ga)
                            sda_unit += 1

            else:
                # Traitement de(s) unité(s)
                nb_unit = plage_sda % 10
                for _ in range(nb_unit):
                    update_list(s_sda, ndi, ga)
                    s_sda += 1

                # Traitement de(s) dixaine(s)
                nb_dix = int(((plage_sda - nb_unit) / 10) % 10)
                if nb_dix >= 1:
                    s_sda_dix = s_sda // 10
                    for _ in range(nb_dix):
                        update_list(s_sda_dix, ndi, ga)
                        s_sda_dix += 1

                    # Traitement de(s) centaine(s)
                    nb_cent = plage_sda // 100
                    if nb_cent >= 1:
                        s_sda_cent = s_sda_dix // 10
                        for _ in range(nb_cent):
                            update_list(s_sda_cent, ndi, ga)
                            s_sda_cent += 1

        else:
            # Traitement si le nombre de SDA est inférieur à 100
            nb_dix = plage_sda // 10
            rest_dix = int(plage_sda % 10)
            if plage_sda >= 10:
                s_sda_dix = str(s_sda)
                s_sda_dix = int(s_sda_dix[:-1])
                for _ in range(nb_dix):
                    update_list(s_sda_dix, ndi, ga)
                    s_sda_dix += 1
                if rest_dix < 10:
                    s_sda_dix *= 10
                    s_sda = s_sda_dix
                    for _ in range(rest_dix):
                        update_list(s_sda, ndi, ga)
                        s_sda += 1

            else:
                for _ in range(plage_sda):
                    update_list(s_sda, ndi, ga)
                    s_sda += 1

    # création d'un dossier sauf si il existe déjà
    os.makedirs('imports/', exist_ok=True)
    print("Création des fichers d'imports de masse dans le dossier imports: En cours ...")

    # Création des fichiers d'import de masse
    df_10D = pd.DataFrame({"OPERATION": "Add", "PUBID": LISTE_SDA, "SED": LISTE_GA})
    df_10D.to_csv('imports/import_routage_SDA.csv', index=False)

    df_ndi = pd.DataFrame({"OPERATION": "Add", "PUBID": LISTE_SDA, "SED": LISTE_NDI})
    df_ndi.to_csv('imports/import_NDI_Frais.csv', index=False)

    print("Création des fichers d'imports de masse dans le dossier imports: Terminé")

else:
    print("Le fichier n'existe pas ou le nom du fichier n'est pas correct.")
    print("Nom du fichier: exemple.xlsx")
    print("Merci de vérifier le nom du fichier ou la présence du fichier dans le même répertoire que l'exécutable.")
os.system("pause")
