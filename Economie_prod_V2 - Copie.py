import pandas as pd
import xlsxwriter
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import matplotlib.pyplot as plt

# Dictionnaire des catégories (inchangé)
categories = {
    'RETRAIT': 'Retraits',
    'HETM': 'Enfants',
    'HELIUM': 'Remboursements',
    'Stand Prive Aulnay': 'Achats & Shopping',
    'DIALOGUES': 'Achats & Shopping',
    'KASHGAR': 'Alimentation & restaurants',
    'GRAND FRAIS': 'Alimentation & restaurants',
    'CHEQUE': 'Santé',
    'Action': 'Achats & Shopping',
    'C.P.A.M.': 'Remboursements',
    'DIRECTION GENERALE DES FINANCES': 'Impôts, taxes & frais',
    'AMAZON': 'Achats & Shopping',
    'DE CHATO JANO': 'Autres revenus',
    'LECLERC': 'Alimentation & restaurants',
    'WWW.PLANITY.COM': 'Enfants',
    'Spotify': 'Loisirs & vacances',
    'KIABI': 'Enfants',
    'SERVICE PUBLIC EAU': 'Logement & charges',
    'MR BRICOLAGE': 'Logement & charges',
    'NETFLIX': 'Loisirs & vacances',
    'SECURIMUT GEM MACIF': 'Crédits',
    'BOUTIQUE SOSH': 'Logement & charges',
    'REMBOURSEMENT DE PRET': 'Crédits',
    'DECATHLON': 'Enfants',
    'BOULANG VIGNARD': 'Alimentation & restaurants',
    'Vinted': 'Achats & Shopping',
    'SNCF': 'Transport',
    'MAE': 'Assurance & prévoyance',
    'SFR': 'Logement & charges',
    'APPLE.COM': 'Logement & charges',
    'ALS ACTION LOGEMENT': 'Crédits',
    'TotalEnergies': 'Logement & charges',
    'CRCAM DU FINISTERE PRELEVEMENT PRET TRAVAUX': 'Crédits',
    'Orange bleue': 'Santé',
    'TRES. BREST CH PREL CHMORLAIX': 'Impôts, taxes & frais',
    'CRF': 'Alimentation & restaurants',
    'JEFF DE BRUGES': 'Loisirs & vacances',
    'H&M':'Enfants',
    'TIPI SMDC':'Enfants',
    'WEEZEVENT DIJON':'Loisirs & vacances',
    'GMF ASSURANCES':'Assurance & prévoyance',
    'AUDIBLE.FR':'Loisirs & vacances',
    'ASSU. CAAE PRET':'Crédits',
    'ARUM':'Loisirs & vacances',
    'PROTECTION JURIDIQUE':'Assurance & prévoyance',
    'Selecta':'Loisirs & vacances',
    'SARL MAD IN BREIZH':'Alimentation & restaurants',
    'PARLUX':'Achats & Shopping',
    'LA POSTE':'Achats & Shopping',
    'ROXYNAILSPARIS.COM':'Loisirs & vacances',
    'BOULANGERIE':'Alimentation & restaurants',
    'FRANCK PROVOST':'Loisirs & vacances',
    'POKE AND CO':'Alimentation & restaurants',
    'ASSOC':'Loisirs & vacances',
    'Orange':'Logement & charges',
    'MALESHERBES':'Loisirs & vacances',
    'Lidl SNC':'Achats & Shopping',
    'AVOIR':'Remboursements',
    'NATURA':'Alimentation & restaurants',
    'PHARMACIE':'Santé',
    'VETERINAIR':'Santé',
    'SUPER U':'Alimentation & restaurants',
    'KEOLIS':'Transport',
    'LA TERRASSE':'Loisirs & vacances',
    'DE FETE':'Loisirs & vacances',
    'POUTINEBROS':'Loisirs & vacances',
    'VINS LAUNAY':'Achats & Shopping',
    'RELAIS PORZ':'Transport',
    'FOURNIL':'Alimentation & restaurants',
    'FEU VERT':'Transport',
    'PñDIATRE':'Santé',
    'CELSOL':'Alimentation & restaurants',
    'Norton':'Assurance & prévoyance',
    'LA PLAYCE':'Loisirs & vacances',
    'LORANGE BLEUE':'Santé',
    'CASINO BENODET':'Loisirs & vacances',
    'TIPI SMDC':'Enfants',
    'INTERMARCHE':'Alimentation & restaurants',
    'MSF MEDECINS':'Impôts, taxes & frais',
    'Intérets débiteurs':'Impôts, taxes & frais',
    'REMBOURSEMENT':'Remboursements',
    'SOCIETE GAZ':'Logement & charges',
    'SHELL':'Transport',
    'PARTS SOC':'Remboursements',
    'Offre Compte':'Impôts, taxes & frais',
    'SHOP-EHCONSULTING':'Loisirs & vacances',
    'Sherlock Pub':'Alimentation & restaurants',
    'SumUp':'Loisirs & vacances',
    'MAXICOFFEE':'Loisirs & vacances',
    'LW-PASSAGEBLEU':'Enfants',
    'PROTEIN':'Alimentation & restaurants',
    'LE DRESSING':'Achats & Shopping',
    'AQUA WEST':'Loisirs & vacances',
    'BURGER84':'Alimentation & restaurants',
    'LASER TAG':'Loisirs & vacances',
    'Interflora':'Achats & Shopping',
    'BLEU LIBELLULE':'Enfants',
    'HPY*DEGUISE-TOI':'Loisirs & vacances',
    'LIDL':'Alimentation & restaurants',
    'BURGER KING':'Alimentation & restaurants',
    'DR LUC DUBRULLE':'Santé',
    'RAKUTEN':'Achats & Shopping',
    'FNAC':'Achats & Shopping',
    'CAMPING':'Loisirs & vacances',
    'DOMAINE D':'Loisirs & vacances',
    'TOTAL':'Transport',
    'PRELEVEMENT C.G.O.S':'Loisirs & vacances',
    'LA TABLE':'Alimentation & restaurants',
    'RELAIS':'Transport',
    'UEP*DAC':'Loisirs & vacances',
    'REL PAYS':'Loisirs & vacances',
    'LES DELICES':'Alimentation & restaurants',
    'LARNAS':'Alimentation & restaurants',
    'LA SERRE':'Loisirs & vacances',
    'PAYE DU MOIS':'Salaire',
    'DIOT':'Salaire',
    'VIREMENT EN VOTRE FAVEUR':'Remboursements',
    'DECATHLON':'Enfants',
    'VIREMENT': 'Virement'
}

# Montants attribués par catégorie
montants_attribues = {
    'Alimentation & restaurants': 1000,
    # Ajoutez d'autres catégories et montants ici si nécessaire
}

def charger_et_preparer_donnees(fichier_entree):
    # Charger les premières lignes pour identifier où commencent les données
    df_preview = pd.read_excel(fichier_entree, nrows=20)
    
    # Trouver la ligne qui contient 'Date' (ou un autre en-tête attendu)
    header_row = df_preview[df_preview.apply(lambda row: 'Date' in row.values, axis=1)].index
    if len(header_row) == 0:
        print("Impossible de trouver la ligne d'en-tête.")
        return None

    # Recharger le fichier en utilisant la ligne d'en-tête trouvée
    df = pd.read_excel(fichier_entree, header=header_row[0])

    # Supprimer la première ligne qui contient les en-têtes dupliqués
    df = df.iloc[1:]

    # Réinitialiser l'index après avoir supprimé la première ligne
    df = df.reset_index(drop=True)

    # Renommer les colonnes si nécessaire
    df = df.rename(columns={
        df.columns[0]: 'Date',
        df.columns[1]: 'Libellé',
        df.columns[2]: 'Débit euros',
        df.columns[3]: 'Crédit euros'
    })

    print("DataFrame chargé et préparé:")
    print(df.head())
    print(df.columns)

    return df

def categoriser_transaction(libelle):
    for mot_cle, categorie in categories.items():
        if mot_cle.lower() in libelle.lower():
            return categorie
    return 'Autres'

def selectionner_fichier():
    root = tk.Tk()
    root.withdraw()
    fichier_selectionne = filedialog.askopenfilename(
        title="Sélectionner un fichier Excel",
        filetypes=[("Fichiers Excel", "*.xlsx *.xls")]
    )
    return fichier_selectionne

def selectionner_repertoire():
    root = tk.Tk()
    root.withdraw()
    repertoire_selectionne = filedialog.askdirectory(
        title="Sélectionner un répertoire de sortie"
    )
    return repertoire_selectionne

def creer_camemberts(df, worksheet_depenses, worksheet_revenus):
    # Calculer les dépenses et revenus par catégorie
    depenses = df[df['Débit euros'] > 0].groupby('Catégorie')['Débit euros'].sum()
    revenus = df[df['Crédit euros'] > 0].groupby('Catégorie')['Crédit euros'].sum()

    # Créer le camembert des dépenses
    fig_depenses, ax_depenses = plt.subplots(figsize=(8, 6))
    ax_depenses.pie(depenses.values, labels=depenses.index, autopct='%1.1f%%', startangle=90)
    ax_depenses.set_title('Répartition des dépenses par catégorie')
    plt.tight_layout()

    # Sauvegarder le camembert des dépenses
    depenses_path = 'depenses_camembert.png'
    fig_depenses.savefig(depenses_path)

    # Insérer le camembert des dépenses dans l'onglet Dépenses
    worksheet_depenses.insert_image('F2', depenses_path)

    # Créer le camembert des revenus
    fig_revenus, ax_revenus = plt.subplots(figsize=(8, 6))
    ax_revenus.pie(revenus.values, labels=revenus.index, autopct='%1.1f%%', startangle=90)
    ax_revenus.set_title('Répartition des revenus par catégorie')
    plt.tight_layout()

    # Sauvegarder le camembert des revenus
    revenus_path = 'revenus_camembert.png'
    fig_revenus.savefig(revenus_path)

    # Insérer le camembert des revenus dans l'onglet Revenus
    worksheet_revenus.insert_image('F2', revenus_path)

    # Fermer les figures pour libérer la mémoire
    plt.close(fig_depenses)
    plt.close(fig_revenus)

def formater_excel(df, fichier_sortie):
    workbook = xlsxwriter.Workbook(fichier_sortie)
    worksheet_depenses = workbook.add_worksheet("Dépenses")
    worksheet_revenus = workbook.add_worksheet("Revenus")

    # Définir les formats
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'align': 'center', 'valign': 'vcenter'})
    libelle_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    montant_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'center', 'valign': 'vcenter'})
    categorie_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

    # Définir les largeurs de colonnes
    worksheet_depenses.set_column('A:A', 15)  # Date
    worksheet_depenses.set_column('B:B', 31)  # Libellé
    worksheet_depenses.set_column('C:C', 15)  # Débit euros
    worksheet_depenses.set_column('D:D', 20)  # Catégorie
    worksheet_depenses.set_column('E:E', 20)  # Montant attribué

    worksheet_revenus.set_column('A:A', 15)  # Date
    worksheet_revenus.set_column('B:B', 31)  # Libellé
    worksheet_revenus.set_column('C:C', 15)  # Crédit euros
    worksheet_revenus.set_column('D:D', 20)  # Catégorie
    worksheet_revenus.set_column('E:E', 20)  # Montant attribué

    # Écrire les en-têtes
    headers = ['Date', 'Libellé', 'Montant', 'Catégorie', 'Montant attribué']
    for col, header in enumerate(headers):
        worksheet_depenses.write(0, col, header, header_format)
        worksheet_revenus.write(0, col, header, header_format)

    # Séparer les dépenses et les revenus
    depenses = df[df['Débit euros'] > 0].copy()  # Utiliser .copy() pour éviter les SettingWithCopyWarning
    revenus = df[df['Crédit euros'] > 0].copy()  # Utiliser .copy() pour éviter les SettingWithCopyWarning

    # Ajouter la colonne "Montant attribué"
    depenses['Montant attribué'] = depenses['Catégorie'].map(montants_attribues).fillna('N/A')
    revenus['Montant attribué'] = revenus['Catégorie'].map(montants_attribues).fillna('N/A')

        # Écrire les données de dépenses
    for row, data in depenses.iterrows():
        worksheet_depenses.write(row + 1, 0, data['Date'], date_format)
        worksheet_depenses.write(row + 1, 1, data['Libellé'], libelle_format)
        worksheet_depenses.write_number(row + 1, 2, data['Débit euros'], montant_format)
        worksheet_depenses.write(row + 1, 3, data['Catégorie'], categorie_format)
        worksheet_depenses.write(row + 1, 4, data['Montant attribué'], categorie_format)

    # Écrire les données de revenus
    for row, data in revenus.iterrows():
        worksheet_revenus.write(row + 1, 0, data['Date'], date_format)
        worksheet_revenus.write(row + 1, 1, data['Libellé'], libelle_format)
        worksheet_revenus.write_number(row + 1, 2, data['Crédit euros'], montant_format)
        worksheet_revenus.write(row + 1, 3, data['Catégorie'], categorie_format)
        worksheet_revenus.write(row + 1, 4, data['Montant attribué'], categorie_format)

    # Ajouter les camemberts
    creer_camemberts(df, worksheet_depenses, worksheet_revenus)
    
    # Créer une nouvelle feuille pour les revenus et dépenses par catégorie
    worksheet_totals = workbook.add_worksheet("Revenus & Dépenses")

    # Calculer les dépenses totales par catégorie
    depenses_par_categorie = df.groupby('Catégorie')['Débit euros'].sum().reset_index()
    depenses_par_categorie = depenses_par_categorie[depenses_par_categorie['Débit euros'] > 0]

    # Calculer les revenus totaux par catégorie
    revenus_par_categorie = df.groupby('Catégorie')['Crédit euros'].sum().reset_index()
    revenus_par_categorie = revenus_par_categorie[revenus_par_categorie['Crédit euros'] > 0]

    # Définir les largeurs de colonnes pour la nouvelle feuille
    worksheet_totals.set_column('A:A', 31)  # Catégorie
    worksheet_totals.set_column('B:B', 15)  # Montant

    # Écrire les en-têtes pour la nouvelle feuille
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    worksheet_totals.write(0, 0, 'Catégorie', header_format)
    worksheet_totals.write(0, 1, 'Montant', header_format)
    worksheet_totals.write(0, 2, 'Type', header_format)

    workbook.close()

# Sélectionner le fichier Excel d'entrée
fichier_entree = selectionner_fichier()
if not fichier_entree:
    print("Aucun fichier sélectionné.")
    exit()

# Charger et préparer les données
df = charger_et_preparer_donnees(fichier_entree)
if df is None:
    print("Impossible de charger les données correctement.")
    exit()

# Convertir la colonne 'Date' au format souhaité
df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')

# Ajouter la colonne Catégorie
df['Catégorie'] = df['Libellé'].apply(categoriser_transaction)

# Sélectionner le répertoire de sortie
repertoire_sortie = selectionner_repertoire()
if not repertoire_sortie:
    print("Aucun répertoire sélectionné.")
    exit()

# Générer le nom du fichier de sortie avec la date du jour
date_du_jour = datetime.now().strftime('%Y%m%d')
fichier_sortie = f'{repertoire_sortie}/CA_{date_du_jour}.xlsx'

# Formater et sauvegarder le fichier Excel
formater_excel(df, fichier_sortie)

print(f"Fichier catégorisé et formaté sauvegardé sous {fichier_sortie}")