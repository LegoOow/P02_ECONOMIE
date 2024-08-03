import pandas as pd
import xlsxwriter
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import matplotlib.pyplot as plt

# Montants attribués par catégorie
montants_attribues = {
    'Alimentation & restaurants': 11400,
    'Loisirs & vacances': 6727,
    'Achats & Shopping': 6709,
    'Enfants': 0,
    'Santé': 0,
    'Retraits': 0,
    'Impôts, taxes & frais': 0,
    'Logement & charges': 0,
    'Crédits': 0,
    'Transport': 0,
    'Assurance & prévoyance': 0,
    'Salaire': 0,
    # Ajoutez d'autres catégories et montants ici si nécessaire
}

# Dictionnaire des catégories (inchangé)
categories = {
    'RETRAIT': 'Retraits',
    'SAS HUB':'Achats & Shopping',
    '6715190':'Enfants',
    'TRAON':'Remboursements',
    'COULOIGNER':'Remboursements',
    'COZIEN':'Remboursements',
    'DERRIEN':'Remboursements',
    'GRALL':'Remboursements',
    'MARGOT':'Remboursements',
    'CORRE':'Remboursements',
    'SIAC':'Remboursements',
    'DIMITRI':'Remboursements',
    'DIDIER':'Remboursements',
    'OCEANE':'Remboursements',
    'ESPECES':'Remboursements',
    'REMISE DE CHEQUE':'Remboursements',
    'JOKO':'Autres revenus',
    'AMUNDI':'PRIMES',
    'HETM': 'Enfants',
    'HELIUM': 'Remboursements',
    'Stand Prive Aulnay': 'Achats & Shopping',
    'DIALOGUES': 'Achats & Shopping',
    'KASHGAR': 'Alimentation & restaurants',
    'GRAND FRAIS': 'Alimentation & restaurants',
    'AVANCE CREDIMPOT':'Remboursements',
    '18128': 'Santé',
    '4472': 'Achats & Shopping',
    'C.P.A.M.': 'Remboursements',
    'TONTON':'Alimentation & restaurants',
    'RESTO':'Alimentation & restaurants',
    'DIRECTION GENERALE DES FINANCES': 'Impôts, taxes & frais',
    'AMAZON': 'Achats & Shopping',
    'DE CHATO JANO': 'Autres revenus',
    'E.LECLERC': 'Alimentation & restaurants',
    'LECLERC MORLAIX':'Alimentation & restaurants',
    'CENTRE LECLERC':'Alimentation & restaurants',
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
    'ALS ACTION': 'Crédits',
    'TotalEnergies': 'Logement & charges',
    'CRCAM DU FINISTERE PRELEVEMENT PRET TRAVAUX': 'Crédits',
    'Orange bleue': 'Loisirs & vacances',
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
    'FRANCK PROVOST':'Achats & Shopping',
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
    'KEOLIS':'Loisirs & vacances',
    'LA TERRASSE':'Loisirs & vacances',
    'DE FETE':'Loisirs & vacances',
    'POUTINEBROS':'Loisirs & vacances',
    'VINS LAUNAY':'Achats & Shopping',
    'RELAIS PORZ':'Loisirs & vacances',
    'FOURNIL':'Alimentation & restaurants',
    'FEU VERT':'Transport',
    'PñDIATRE':'Santé',
    'CELSOL':'Alimentation & restaurants',
    'Norton':'Assurance & prévoyance',
    'LA PLAYCE':'Loisirs & vacances',
    'LORANGE BLEUE':'Loisirs & vacances',
    'CASINO BENODET':'Loisirs & vacances',
    'TIPI SMDC':'Enfants',
    'INTERMARCHE':'Alimentation & restaurants',
    'MSF MEDECINS':'Impôts, taxes & frais',
    'Intérets débiteurs':'Impôts, taxes & frais',
    'REMBOURSEMENT':'Remboursements',
    'SOCIETE GAZ':'Logement & charges',
    'SHELL':'Loisirs & vacances',
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
    'TOTAL':'Loisirs & vacances',
    'PRELEVEMENT C.G.O.S':'Loisirs & vacances',
    'LA TABLE':'Alimentation & restaurants',
    'RELAIS':'Loisirs & vacances',
    'UEP*DAC':'Loisirs & vacances',
    'REL PAYS':'Loisirs & vacances',
    'LES DELICES':'Alimentation & restaurants',
    'LARNAS':'Alimentation & restaurants',
    'LA SERRE':'Loisirs & vacances',
    'PAYE DU MOIS':'Salaire',
    'DIOT':'Salaire',
    'HELIUM':'Remboursements',
    'CASTEL KEVIN':'Virement',
    'CAF':'Aide Sociale',
    'DECATHLON':'Enfants',
    'BOUYGUES':'Logement & charges',
    'Rue ar Brug':'Achats & Shopping',
    'YEPODA.FR':'Achats & Shopping',
    'LA COMBE':'Loisirs & vacances',
    'POKEFACT':'Alimentation & restaurants',
    'PASSAGE BLEU':'Enfants',
    'C.G.O.S':'Loisirs & vacances',
    'GROTTE CHAUVET':'Loisirs & vacances',
    'UEP*PLEYBERIENNE':'Loisirs & vacances',
    'La Ferme Aux Crocodi':'Loisirs & vacances',
    'DAC ANGERS':'Loisirs & vacances',
    'ALLIER DOYET':'Loisirs & vacances',
    'LIBRAIRES':'Loisirs & vacances',
    'LECLERCSTATION':'Transport',
    'REL LES VOLCANS':'Loisirs & vacances',
    'AVIA':'Loisirs & vacances',
    'MONOPRIX':'Alimentation & restaurants',
    '6715207':'Santé',
    'BAKER ROSS':'Achats & Shopping',
    'CASTORAMA':'Achats & Shopping',
    'MORD L EXPRESS':'Alimentation & restaurants',
    'SEW':'Loisirs & vacances',
    'AMENDE.GOUV':'Transport',
    'UEP*SODEXLAND':'Transport',
    'ESSFLOREAL':'Transport',
    'SIM-SUSHIS':'Alimentation & restaurants',
    'PICARD':'Alimentation & restaurants',
    'JOUECLUB':'Achats & Shopping',
    'FDJ BOULOGNE':'Loisirs & vacances',
    'NATURE ALIMENTS':'Alimentation & restaurants',
    'CMC DE LA BAIE':'Santé',
    'GIBIER LAURE':'Santé',
    'MAC DONALD':'Alimentation & restaurants',
    'COFIROUTE':'Loisirs & vacances',
    'APS PARIS':'Loisirs & vacances',
    'GARAGE DU PUITS':'Loisirs & vacances',
    'GEANT':'Alimentation & restaurants',
    'JDCB':'Alimentation & restaurants',
    'LE POT COMMUN':'Achats & Shopping',
    'DOMAINE IMBOURS':'Loisirs & vacances',
    'ATELIER DES PAINS':'Alimentation & restaurants',
    'DISTRIVERT':'Logement & charges',
    'OCIE':'Transport',
    'POUVOIR DES FLEURS':'Achats & Shopping',
    'Nike':'Achats & Shopping',
    'ANIMOTOPIA':'Alimentation & restaurants',
    'Clinique France':'Achats & Shopping',
    'INTERSPORT':'Achats & Shopping',
    'MAUSSION ELOISE':'Santé',
    'LERIALTO':'Loisirs & vacances',
    'NORMAL':'Alimentation & restaurants',
    'MARY STUART':'Alimentation & restaurants',
    'RIALTO':'Loisirs & vacances',
    'AURORE':'Loisirs & vacances',
    'Abritel':'Loisirs & vacances',
    'PEAGE':'Loisirs & vacances',
    'HELLOBODY':'Achats & Shopping',
    'ASF':'Loisirs & vacances',
    'CRCAM DU FINISTERE':'Crédits',
    'CARREFOUR':'Loisirs & vacances',
    'STATION HYPER':'Loisirs & vacances',
    'REL LES VOLCANS':'Loisirs & vacances',
    'YFFINIAC':'Loisirs & vacances',
    'LE SLALOM':'Loisirs & vacances',
    'SAS VILLANS':'Loisirs & vacances',
    'SSP SAINT-GILLES':'Loisirs & vacances',
    'NLS VERRIERES':'Loisirs & vacances',
    'ALTITUDE 2000':'Loisirs & vacances',
    'CHORANCHE':'Loisirs & vacances',
    'VILLANS':'Loisirs & vacances',
    'VILLARD':'Loisirs & vacances',
    'RM ESPACE':'Loisirs & vacances',
    'A.R.E.A.':'Loisirs & vacances',
    'GUILLET':'Loisirs & vacances',
    'APRR':'Loisirs & vacances',
    'EKINSPOR':'Loisirs & vacances',
    'MALMAISON':'Loisirs & vacances',
    'FARGES':'Loisirs & vacances',
    'NLS DRUYE':'Loisirs & vacances',
    'BPNL':'Loisirs & vacances',
    'DRONIOU':'Transport',
    'GEMO':'Enfants',
    'FERNAND':'Alimentation & restaurants',
    'NESTLE':'Alimentation & restaurants',
    'FBGN':'Alimentation & restaurants',
    'MGP*Reelax':'Loisirs & vacances',
    'STATIONNEMENTMLX':'Transport',
    'ALEXANDRA':'Santé',
    'DISTRI':'Achats & Shopping',
    'CHATELAIN':'Alimentation & restaurants',
    'VEOLIA':'Logement & charges',
    'DR SEIGNEURIC':'Santé',
    'TEMU.COM':'Achats & Shopping',
    'LOPIN':'Santé',
    'SPAR':'Alimentation & restaurants',
    'MEROUR':'Alimentation & restaurants',
    'BIJOU':'Achats & Shopping',
    'eBay':'Achats & Shopping',
    'BURGER':'Alimentation & restaurants',
    'MEYL':'Loisirs & vacances',
    'JUNGLE':'Loisirs & vacances',
    'ANDRE':'Enfants',
    'CINEMA':'Loisirs & vacances',
    'Le Palais':'Alimentation & restaurants',
    'FROMAGES':'Alimentation & restaurants',
    'VF J':'Achats & Shopping',
    'PPG*VETSECURITE':'Achats & Shopping',
    'TY COZ':'Loisirs & vacances',
    'DR SEGUIN':'Santé',
    'MAD DO':'Alimentation & restaurants',
    '6715201':'Santé',
    'AU BUREAU':'Alimentation & restaurants',
    'CHATGPT':'Loisirs & vacances',
    'TREVAREZ':'Loisirs & vacances',
    'CLIMB UP':'Loisirs & vacances',
    'ATAORIDE':'Loisirs & vacances',
    'CAPU':'Loisirs & vacances',
    'DEFOUL':'Loisirs & vacances',
    '6715203':'Logement & charges',
    'MEDIATHEQUE':'Enfants',
    'IMAGERIE':'Santé',
    'DR LE PAGE':'Santé',
    'ILOISE':'Alimentation & restaurants',
    'MEMPHIS':'Alimentation & restaurants',
    'MERLIN':'Logement & charges',
    'KaffeKapslen':'Alimentation & restaurants',
    'BIANNIC':'Alimentation & restaurants',
    'Back Market':'Achats & Shopping',
    'LEBON':'Achats & Shopping',
    'IDKIDS':'Enfants',
    'CONCARNEAU':'Loisirs & vacances',
    'STOKOMANI':'Achats & Shopping',
    'WWW.SMYTHSTOYS.COM':'Achats & Shopping',
    'OBERNAI':'Loisirs & vacances',
    'PISCINE':'Loisirs & vacances',
    'TYSHALA':'Loisirs & vacances',
    'PAYLINE':'Achats & Shopping',
    'SHOWROOMPR':'Achats & Shopping',
    'LE LOFT':'Enfants',
    'LES BAINS':'Alimentation & restaurants',
    'MAGIC':'Loisirs & vacances',
    'WONDERBOX':'Achats & Shopping',
    'LS MORLAIX':'Loisirs & vacances',
    'BOUL P':'Alimentation & restaurants',
    'H M Hennes':'Achats & Shopping',
    'TEMPO':'Loisirs & vacances',
    'MOULIN':'Alimentation & restaurants',
    '6715194':'Loisirs & vacances',
    '6715193':'Loisirs & vacances',
    'SENSAS':'Loisirs & vacances',
    'VERTBAUDET.FR':'Achats & Shopping',
    '6715197/0000000/000000000':'Santé',
    'KING JOUET':'Achats & Shopping',
    'BONJOUR MINUIT':'Loisirs & vacances',
    'FETE-CI':'Loisirs & vacances',
    'HIPPOPOTAMUS':'Alimentation & restaurants',
    '6715196':'Achats & Shopping',
    'IKEA':'Logement & charges',
    'CDISCOUNT':'Achats & Shopping',
    'BUREAU VALLEE':'Achats & Shopping',
    'Sony Interactive':'Loisirs & vacances',
    'DR LE CLAIRE':'Santé',
    'LE CORSAIRE':'Alimentation & restaurants',
    '6715192/0000000/000000000':'Loisirs & vacances',
    'ESPRIT FAMILLE':'Alimentation & restaurants',
    '6715208':'Santé',
    'Free Mobile':'Achats & Shopping',
    'OPTICAL':'Santé',
    'CFV TOURNON':'Loisirs & vacances',
    'TOUR EIFFEL':'Loisirs & vacances',
    '6715195/0000000/000000000':'Santé',
    'PAYGREEN':'Loisirs & vacances',
    'SPA NANCY':'Loisirs & vacances',
    'LA CHOCOLATERI':'Achats & Shopping',
    'O QUATORZE':'Alimentation & restaurants',
    'Weekendesk':'Loisirs & vacances',
    '6715191':'Santé',
    'LES ALIZES':'Alimentation & restaurants',
    'WEEZEVENT':'Loisirs & vacances',
    'CARHARTT':'Achats & Shopping',
    'SEPHORA':'Achats & Shopping',
    'ETIQUETTES':'Enfants',
    'CASTEL MAHO': 'Epargne',
    'CASTEL NOE':'Epargne',
    'Ma Banque':'Epargne',
    'Fortuneo':'Loisirs & vacances',
    'PENNANEACH':'Remboursements',
    'Paylib':'Remboursements',
    'WEB Castel Dimitri':'Logement & charges',
    'Fabien Couloignier':'Remboursements',
    'ECHAPPEE':'Loisirs & vacances',
    'Fourniture carte débit':'Logement & charges',
    'HOTEL':'Loisirs & vacances',
    'LA PETRISSEE':'Loisirs & vacances',
    'DONALD':'Alimentation & restaurants',
    'ZOO':'Loisirs & vacances',
    'EasyPark':'Loisirs & vacances',
    'ENEDIS':'Logement & charges',
    'BRICO':'Logement & charges',
    'MEDECINS SANS FRONTIERES':'Impôts, taxes & frais',
    'PHILIBERTNET.COM':'Enfants',
    'LA HALLE':'Enfants',
    'Motocultor festival':'Loisirs & vacances'
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
    # Calculer les dépenses et revenus par catégorie, en excluant "Virement"
    depenses = df[(df['Débit euros'] > 0) & (df['Catégorie'] != 'Virement')].groupby('Catégorie')['Débit euros'].sum()
    revenus = df[(df['Crédit euros'] > 0) & (df['Catégorie'] != 'Virement')].groupby('Catégorie')['Crédit euros'].sum()

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
    worksheet_recap = workbook.add_worksheet("Récapitulatif")
    worksheet_totals = workbook.add_worksheet("Revenus & Dépenses")

    # Définir les formats
    header_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'align': 'center', 'valign': 'vcenter'})
    libelle_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    montant_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'center', 'valign': 'vcenter'})
    categorie_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

    # Définir les largeurs de colonnes pour les onglets Dépenses et Revenus
    for worksheet in [worksheet_depenses, worksheet_revenus]:
        worksheet.set_column('A:A', 15)  # Date
        worksheet.set_column('B:B', 31)  # Libellé
        worksheet.set_column('C:C', 15)  # Montant
        worksheet.set_column('D:D', 20)  # Catégorie

    # Écrire les en-têtes pour les onglets Dépenses et Revenus
    headers = ['Date', 'Libellé', 'Montant', 'Catégorie']
    for worksheet in [worksheet_depenses, worksheet_revenus]:
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

    # Séparer les dépenses et les revenus
    depenses = df[df['Débit euros'] > 0].copy()
    revenus = df[df['Crédit euros'] > 0].copy()

    # Écrire les données de dépenses
    for row, data in depenses.iterrows():
        worksheet_depenses.write(row + 1, 0, data['Date'], date_format)
        worksheet_depenses.write(row + 1, 1, data['Libellé'], libelle_format)
        worksheet_depenses.write_number(row + 1, 2, data['Débit euros'], montant_format)
        worksheet_depenses.write(row + 1, 3, data['Catégorie'], categorie_format)

    # Écrire les données de revenus
    for row, data in revenus.iterrows():
        worksheet_revenus.write(row + 1, 0, data['Date'], date_format)
        worksheet_revenus.write(row + 1, 1, data['Libellé'], libelle_format)
        worksheet_revenus.write_number(row + 1, 2, data['Crédit euros'], montant_format)
        worksheet_revenus.write(row + 1, 3, data['Catégorie'], categorie_format)

    # Créer les tableaux récapitulatifs
    depenses_par_categorie = df.groupby('Catégorie')['Débit euros'].sum().reset_index()
    revenus_par_categorie = df.groupby('Catégorie')['Crédit euros'].sum().reset_index()

    # Écrire les tableaux récapitulatifs dans l'onglet Récapitulatif
    worksheet_recap.write(0, 0, "Dépenses par catégorie", header_format)
    worksheet_recap.write(0, 1, "Montant", header_format)
    for row, data in depenses_par_categorie.iterrows():
        worksheet_recap.write(row + 1, 0, data['Catégorie'], categorie_format)
        worksheet_recap.write_number(row + 1, 1, data['Débit euros'], montant_format)

    worksheet_recap.write(0, 3, "Revenus par catégorie", header_format)
    worksheet_recap.write(0, 4, "Montant", header_format)
    for row, data in revenus_par_categorie.iterrows():
        worksheet_recap.write(row + 1, 3, data['Catégorie'], categorie_format)
        worksheet_recap.write_number(row + 1, 4, data['Crédit euros'], montant_format)

    # Écrire le tableau des dépenses avec montant attribué dans l'onglet Revenus & Dépenses
    worksheet_totals.write(0, 0, 'Catégorie', header_format)
    worksheet_totals.write(0, 1, 'Montant', header_format)
    worksheet_totals.write(0, 2, 'Montant attribué', header_format)

    for row, data in depenses_par_categorie.iterrows():
        worksheet_totals.write(row + 1, 0, data['Catégorie'], categorie_format)
        worksheet_totals.write_number(row + 1, 1, data['Débit euros'], montant_format)
        montant_attribue = montants_attribues.get(data['Catégorie'], 'N/A')
        worksheet_totals.write(row + 1, 2, montant_attribue, categorie_format)

    # Ajouter les camemberts
    creer_camemberts(df, worksheet_depenses, worksheet_revenus)

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
    worksheet_recap = workbook.add_worksheet("Récapitulatif")
    worksheet_totals = workbook.add_worksheet("Revenus & Dépenses")

    # Définir les formats
    header_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'align': 'center', 'valign': 'vcenter'})
    libelle_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    montant_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'center', 'valign': 'vcenter'})
    categorie_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

    # Définir les largeurs de colonnes pour les onglets Dépenses et Revenus
    for worksheet in [worksheet_depenses, worksheet_revenus]:
        worksheet.set_column('A:A', 15)  # Date
        worksheet.set_column('B:B', 31)  # Libellé
        worksheet.set_column('C:C', 15)  # Montant
        worksheet.set_column('D:D', 20)  # Catégorie

    # Écrire les en-têtes pour les onglets Dépenses et Revenus
    headers = ['Date', 'Libellé', 'Montant', 'Catégorie']
    for worksheet in [worksheet_depenses, worksheet_revenus]:
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

    # Séparer les dépenses et les revenus
    depenses = df[df['Débit euros'] > 0].copy()
    revenus = df[df['Crédit euros'] > 0].copy()

    # Écrire les données de dépenses
    for row, data in depenses.iterrows():
        worksheet_depenses.write(row + 1, 0, data['Date'], date_format)
        worksheet_depenses.write(row + 1, 1, data['Libellé'], libelle_format)
        worksheet_depenses.write_number(row + 1, 2, data['Débit euros'], montant_format)
        worksheet_depenses.write(row + 1, 3, data['Catégorie'], categorie_format)

    # Écrire les données de revenus
    for row, data in revenus.iterrows():
        worksheet_revenus.write(row + 1, 0, data['Date'], date_format)
        worksheet_revenus.write(row + 1, 1, data['Libellé'], libelle_format)
        worksheet_revenus.write_number(row + 1, 2, data['Crédit euros'], montant_format)
        worksheet_revenus.write(row + 1, 3, data['Catégorie'], categorie_format)

    # Créer les tableaux récapitulatifs
    depenses_par_categorie = df.groupby('Catégorie')['Débit euros'].sum().reset_index()
    revenus_par_categorie = df.groupby('Catégorie')['Crédit euros'].sum().reset_index()

    # Écrire les tableaux récapitulatifs dans l'onglet Récapitulatif
    worksheet_recap.write(0, 0, "Dépenses par catégorie", header_format)
    worksheet_recap.write(0, 1, "Montant", header_format)
    for row, data in depenses_par_categorie.iterrows():
        worksheet_recap.write(row + 1, 0, data['Catégorie'], categorie_format)
        worksheet_recap.write_number(row + 1, 1, data['Débit euros'], montant_format)

    worksheet_recap.write(0, 3, "Revenus par catégorie", header_format)
    worksheet_recap.write(0, 4, "Montant", header_format)
    for row, data in revenus_par_categorie.iterrows():
        worksheet_recap.write(row + 1, 3, data['Catégorie'], categorie_format)
        worksheet_recap.write_number(row + 1, 4, data['Crédit euros'], montant_format)

    # Écrire le tableau des dépenses avec montant attribué dans l'onglet Revenus & Dépenses
    worksheet_totals.write(0, 0, 'Catégorie', header_format)
    worksheet_totals.write(0, 1, 'Montant', header_format)
    worksheet_totals.write(0, 2, 'Montant attribué', header_format)

    for row, data in depenses_par_categorie.iterrows():
        worksheet_totals.write(row + 1, 0, data['Catégorie'], categorie_format)
        worksheet_totals.write_number(row + 1, 1, data['Débit euros'], montant_format)
        montant_attribue = montants_attribues.get(data['Catégorie'], 'N/A')
        worksheet_totals.write(row + 1, 2, montant_attribue, categorie_format)

    # Ajouter les camemberts
    creer_camemberts(df, worksheet_depenses, worksheet_revenus)

    workbook.close()

def ajouter_donnees_resume(classeur, feuille, depenses, revenus):
    # Définir les styles
    header_format = classeur.add_format({
        'bold': True,
        'bg_color': '#DDDDDD',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    cell_format = classeur.add_format({'border': 1})
    total_format = classeur.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Ajouter les en-têtes pour les dépenses
    feuille.write(0, 0, "Catégorie", header_format)
    feuille.write(0, 1, "Dépenses", header_format)
    feuille.write(0, 2, "Montant attribué", header_format)

    # Ajouter les données de dépenses
    ligne = 1
    for _, ligne_dep in depenses.iterrows():
        feuille.write(ligne, 0, ligne_dep['Catégorie'], cell_format)
        feuille.write(ligne, 1, ligne_dep['Débit euros'], cell_format)
        feuille.write(ligne, 2, ligne_dep['Montant attribué'], cell_format)
        ligne += 1

    # Ajouter les en-têtes pour les revenus
    ligne += 2  # Ajouter un espace entre les deux tableaux
    feuille.write(ligne, 0, "Catégorie", header_format)
    feuille.write(ligne, 1, "Revenus", header_format)

    # Ajouter les données de revenus
    ligne += 1
    for _, ligne_rev in revenus.iterrows():
        feuille.write(ligne, 0, ligne_rev['Catégorie'], cell_format)
        feuille.write(ligne, 1, ligne_rev['Crédit euros'], cell_format)
        ligne += 1

    # Calculer et ajouter les totaux
    total_depenses = depenses['Débit euros'].sum()
    total_revenus = revenus['Crédit euros'].sum()
    derniere_ligne_depenses = depenses.shape[0] + 1
    derniere_ligne_revenus = derniere_ligne_depenses + revenus.shape[0] + 3

    feuille.write(derniere_ligne_depenses, 0, "Total Dépenses", total_format)
    feuille.write(derniere_ligne_depenses, 1, total_depenses, total_format)

    feuille.write(derniere_ligne_revenus, 0, "Total Revenus", total_format)
    feuille.write(derniere_ligne_revenus, 1, total_revenus, total_format)

    # Ajuster la largeur des colonnes
    feuille.set_column(0, 0, 20)  # Catégorie
    feuille.set_column(1, 2, 15)  # Montant et Montant attribué

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