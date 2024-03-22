import adjustText
from adjustText import adjust_text
from collections import OrderedDict
import copy
from datetime import datetime
from datetime import timedelta
import exrex
import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.ticker import ScalarFormatter
from mpl_toolkits.axes_grid1 import make_axes_locatable
from mpl_toolkits.axes_grid1.inset_locator import inset_axes
import matplotlib.cm as cm
from matplotlib.colors import Normalize
import mplcursors
import numpy as np
import os.path
import pandas as pd
from pandastable import Table
from pprint import pprint
import re
import seaborn as sns
import shutil
import textwrap
from termcolor import colored
import time
import tkinter as tk
from tkinter import *
import unicodedata
import uuid
import warnings
import xml.etree.ElementTree as ET

def sym_distrib(mySeries, n_std):

    # Défintion des 4 fonctions
    transforms = {'real': mySeries,
                  'sqrt': np.sqrt(mySeries),
                  'cbrt': np.cbrt(mySeries),
                  'log': np.log(mySeries)}
    
    # Calcul skewness (valeur absolue) pour chacune des 4 distributions obtenues
    skews = {key: abs(val.skew()) for key, val in transforms.items()} 
    
    # Transformation avec la skewness minimum (en valeur abs)
    min_skew = min(skews, key=skews.get) 

    # Distribution symétrique, sa moyenne et son écart-type
    mySeries_sym = transforms[min_skew] # f(Serie)
    mean = mySeries_sym.mean() # Moy de f(Serie)
    std = mySeries_sym.std() # Std de f(Serie)

    # Distribution symétrique filtrée (on retire tout ce qui est au-delà de moyenne +- n*sigma)
    mySeries_sym_filtered = mySeries_sym[mySeries_sym >= mean - n_std*std][mySeries_sym <= mean + n_std*std] # f(Series) filtrée (n_std autour de la moy)
    
    # Calcul des bornes sur la distribution avant transformation
    if not(mySeries_sym_filtered.empty):
        inf = mySeries[mySeries_sym_filtered.idxmin()] # Borne inf à partir du min de f(Series) filtrée
        sup = mySeries[mySeries_sym_filtered.idxmax()] 
    else:
        inf = mySeries.min()
        sup = mySeries.max()

    return inf, sup, mean, std, min_skew



























############################################################################################################################
#------------------------------------------ IMPORT ET CORRECTIONS PRELIMINAIRES -------------------------------------------#
############################################################################################################################

def import_corr_prelim(OPERAT_file, CAP_file):

    # Import données OPERAT
    warnings.simplefilter(action="ignore")
    OPERAT = pd.read_excel(OPERAT_file,
                      sheet_name = 0,
                      header = 0,
                      skiprows = 0,
                      dtype = object)
    warnings.resetwarnings()

    # Lignes vides, Apostrophes, espaces multiples, espaces en début et fin de chaînes de caractères
    OPERAT.dropna(axis=0, how="all", inplace=True) 
    OPERAT = OPERAT.map(lambda x: x.replace("’", "'") if isinstance(x, str) else x, na_action="ignore")
    OPERAT = OPERAT.map(lambda x: " ".join(x.split()) if isinstance(x, str) else x, na_action="ignore")
    OPERAT = OPERAT.map(lambda x: x.strip() if isinstance(x, str) else x, na_action="ignore")


    # Import de la liste des catégories et des bornes
    warnings.simplefilter(action="ignore")
    CAP = pd.read_excel(CAP_file,
                      sheet_name = "Catégories & Sous-catégories",
                      header = 0,
                      skiprows = 0,
                      dtype = object)
    warnings.resetwarnings()
    
    CAP = CAP.map(lambda x: x.replace("’", "'") if isinstance(x, str) else x, na_action="ignore")
    CAP = CAP.map(lambda x: " ".join(x.split()) if isinstance(x, str) else x, na_action="ignore")
    CAP = CAP.map(lambda x: x.strip() if isinstance(x, str) else x, na_action="ignore")
    
    # Liste des CAP et SCAP possibles (v5)
    list_CAP = CAP.loc[CAP["Type"] == "Catégorie"]["Catégories et sous-catégories"].value_counts().index.to_list()
    list_SCAP = CAP.loc[CAP["Type"] == "Sous-Catégorie"]["Catégories et sous-catégories"].value_counts().index.to_list()

    #########################################################################
    # Renomage des CAP pour qu'elles correspondent à la dernière version (v5)
    #########################################################################
    CAP_renaming = {"Enseignement secondaire": "Enseignement Secondaire",
                    "Enseignement supérieur": "Enseignement Supérieur",
                    "Gares routières": "Transport urbain de voyageurs (métro)",
                    "Restauration": "Restauration - Débit de boissons",
                    "Résidences de tourisme et de loisirs": "Résidence de tourisme et loisirs",
                    "Santé et action sociale": "Santé - Etablissements médico-sociaux"
    }
    
    col_CAP = ["Catégorie d'activité majoritaire", "Catégorie d'activité secondaire", "Catégorie d'activité principale"]
    col_SCAP = ["Sous-catégorie d'activité majoritaire", "Sous-catégorie d'activité secondaire"]
    
    OPERAT[col_CAP] = OPERAT[col_CAP].replace(CAP_renaming)
    
    ##########################################################################
    # Renomage des SCAP pour qu'elles correspondent à la dernière version (v5)
    ##########################################################################
    
    # Certaines sous-catégories sont en réalité des catégories. Lorsque c'est le cas, on remplace par la sous-catégorie par défaut.
    
    # On renomme d'abord les SCAP (cette fonction ne va s'appliquer qu'aux SCAP qui sont des CAP)
    OPERAT[col_SCAP] = OPERAT[col_SCAP].replace(CAP_renaming)
    
    # Liste des SCAP renseignées par les assujettis dans OPERAT
    list_SCAP_M = OPERAT["Sous-catégorie d'activité majoritaire"].value_counts().index.to_list()
    list_SCAP_S = OPERAT["Sous-catégorie d'activité secondaire"].value_counts().index.to_list()
    list_SCAP_OPERAT = list(set(list_SCAP_M + list_SCAP_S))
    
    # Pour chaque SCAP renseignée qui est dans la liste des CAP
    for scap in [x for x in list_SCAP_OPERAT if x in list_CAP]:
        
        # On récupère le code de la CAP correspondante
        code_CAP = CAP.loc[CAP["Type"] == "Catégorie"].loc[CAP["Catégories et sous-catégories"] == scap]["Code"].iloc[0]
    
        # Puis on récupère la liste des SCAP correspondant à cette CAP
        scap_list = CAP.loc[CAP["Type"] == "Sous-Catégorie"].loc[CAP["Code"].fillna("").str.contains(code_CAP)]["Catégories et sous-catégories"].to_list()
    
        # On sélectionne la SCAP par défaut (si elle existe, sinon, on passe) et on remplace dans les données OPERAT
        default_scap = [x for x in scap_list if 'defaut' in unicodedata.normalize('NFKD', x).encode('ASCII', 'ignore').decode('ASCII').lower()]
        if len(default_scap) == 1:
            
            default_scap = default_scap[0]
        
            idx_m = OPERAT.loc[OPERAT["Sous-catégorie d'activité majoritaire"] == scap].index
            idx_s = OPERAT.loc[OPERAT["Sous-catégorie d'activité secondaire"] == scap].index
            if len(idx_m) != 0:
                OPERAT.loc[idx_m, "Sous-catégorie d'activité majoritaire"] = default_scap
            if len(idx_s) != 0:
                OPERAT.loc[idx_s, "Sous-catégorie d'activité secondaire"] = default_scap
        
        else:
            continue
    
    # Renomage des SCAP pour qu'elles correspondent à la dernière version (v5)
    SCAP_renaming = {"Activité de santé libérale - Kinésithérapie, Rééducation fonctionnelle": "Activité de santé libérale - Kinésithérapie, Rééducation fonctionnelle,",
                     "Aéroport commercial - Hangars de maintenance aéronautique (gros porteurs) - Densité énergétique Niveau 1 (DE = A W/m²)": "Aéroport commercial - Hangar de maintenance aéronautique (gros porteurs) - Densité énergétique Niveau 1 (DE ≤ A W/m²)",
                     "Aéroport commercial - Hangars de maintenance aéronautique (gros porteurs) - Densité énergétique Niveau 3 (DE = B W/m²)": "Aéroport commercial - Hangars de maintenance aéronautique (gros porteurs) - Densité énergétique Niveau 3 (DE ≥ B W/m²)",
                     "Aéroport commercial - Tri bagages - Densité énergétique Niveau 1 (DE = A W/m²)": "Aéroport commercial - Tri bagages - Densité énergétique Niveau 1 (DE ≤ A W/m²)",
                     "Aéroport commercial - Tri bagages - Densité énergétique Niveau 3 (DE = B W/m²)": "Aéroport commercial - Tri bagages - Densité énergétique Niveau 3 (DE ≥ B W/m²)",
                     "Blanchisserie": "Blanchisserie industrielle (Valeur par défaut)",
                     "Collège": "Enseignement Secondaire (Valeur par défaut)",
                     "Commerces et service de détail - Accessoire de mode (Bijouterie, …) - Zone de vente": "Commerce et service de détail - Accessoire de mode (Bijouterie, …) - Zone de vente",
                     "Commerces et service de détail - Bien être (Sauna et Hammam) - Zone Public": "Commerce et service de détail - Bien être (Sauna -Hammam) - Zone Public",
                     "Commerce et service de détail  - Bien être (Sauna -Hammam) - Zone Public": "Commerce et service de détail - Bien être (Sauna -Hammam) - Zone Public",
                     "Commerces et service de détail - Culture, Média et Loisirs (Libraire, produits culturels, jeux et loisirs...) - Zone de vente": "Commerce et service de détail - Culture, Média et Loisirs (Libraire, produits culturels, jeux et loisirs...) - Zone de vente",
                     "Commerces et service de détail - Equipement de la personne (Vêtements, lingeries, linges de maison, chaussures, maroquinerie et bagages...) - Zone de vente": "Commerce et service de détail - Equipement de la personne (Vêtements, lingeries, linges de maison, chaussures, maroquinerie et bagages...) - Zone de vente",
                     "Commerces et service de détail - Equipement de la personne et Loisirs (Valeur par défaut)": "Commerce et service de détail - Equipement de la personne et Loisirs (Valeur par défaut)",
                     "Commerces et service de détail - Equipement de la personne et Loisirs - Administration et bureaux": "Commerce et service de détail - Equipement de la personne et Loisirs - Administration et bureaux",
                     "Commerces et service de détail - Equipement de la personne et Loisirs - Réserve d'approche": "Commerce et service de détail - Equipement de la personne et Loisirs - Réserve d'approche",
                     "Commerces et service de détail - Numérique et téléphonie - Zone vente": "Commerce et service de détail - Numérique et téléphonie - Zone vente",
                     "Commerces et service de détail - Santé Optique - Zone vente": "Commerce et service de détail - Santé Optique - Zone vente",
                     "Commerces et service de détail - Santé, Soins (Pharmacie, Parapharmacie) - Zone vente": "Commerce et service de détail - Santé, Soins (Pharmacie, Parapharmacie) - Zone vente",
                     "Commerces et service de détail - Service Conseil (Agences de voyages…) - Zone vente": "Commerce et service de détail - Service Conseil (Agences de voyages…) - Zone vente",
                     "Commerces et service de détail - Service Laverie automatique - Zone public": "Commerce et service de détail - Service Laverie automatique - Zone public",
                     "Commerces et service de détail - Service Pressing - Accueil public et process": "Commerce et service de détail - Service Pressing - Accueil public et process",
                     "Commerces et service de détail - Services Equipements de la personne (Cordonnerie, Couturier,…) - Accueil public et process": "Commerce et service de détail - Services Equipements de la personne (Cordonnerie, Couturier,…) - Accueil public et process",
                     "Commerces et service de détail - Soins & Beauté (Beauté & bien être) - Zone soins": "Commerce et service de détail - Soins & Beauté (Beauté & bien être) - Zone soins",
                     "Commerces et service de détail - Soins & Beauté (Parfumerie, cosmétique…) - Zone vente": "Commerce et service de détail - Soins & Beauté (Parfumerie, cosmétique…) - Zone vente",
                     "Commerces et service de détail - Soins de la personne (Coiffeur, Salon d'esthétique, Massage) - Zone vente": "Commerce et service de détail - Soins de la personne (Coiffeur, Salon d'esthétique, Massage) - Zone vente",
                     "Commerces et service de détail - Sports et Outdoor - Zone vente": "Commerce et service de détail - Sports et Outdoor - Zone vente",
                     "Enseignement Secondaire - Lycée d'enseignement général et techhnologique agricole (LGTA) - Lycée d'enseignement professionnel agricole (LEP agricole ) - Toutes séries confondues (Valeur Témoin)": "Enseignement Secondaire - Lycée d'enseignement général et technologique agricole (LGTA) - Lycée d'enseignement professionnel agricole (LEP agricole) - Toutes séries confondues (Valeur Témoin)",
                     "Enseignement Supérieur - Ateliers et Halles techniques avec Process - Niveau 1 (DE = A W/m²)": "Enseignement Supérieur - Ateliers et Halles techniques avec Process - Niveau 1 (DE ≤ A W/m²)",
                     "Enseignement Supérieur - Ateliers et Halles techniques avec Process - Niveau 3 (DE = B W/m²)": "Enseignement Supérieur - Ateliers et Halles techniques avec Process - Niveau 3 (DE ≥ B W/m²)",
                     "Gare ferroviaire - Zones de remisage couvertes et closes": "Gare ferroviaire - Zones de remisage couverte et close",
                     "Lycée d'enseignement général": "Enseignement Secondaire - Lycée d'enseignement général (LG) ou Salles d'enseignement banalisé - Toutes séries confondues (Valeur Témoin)",
                     "Lycée d'enseignement général et technologique – Lycée polyvalent": "Enseignement Secondaire - Lycée d'enseignement général et techhnologique (LGT) - Lycée d'enseignement polyvalent (LEP) - Toutes séries confondues (Valeur Témoin)",
                     "Lycée d'enseignement professionnel": "Enseignement Secondaire - Lycée d'enseignement professionnel",
                     "Salle serveur & Data Center (Valeur par défaut)": "Salle serveur & Data Center - Valeur par défaut",
                     "Stockage à température ambiante": "Logistique température ambiante",
                     "Transport urbain (Valeur par défaut)": "Transport urbain - Valeur par défaut",
                     "nseignement Secondaire - Salles de TP - Série SMS/ST2S Sciences médicosocial/ sciences et technologies de la santé et du social": "Enseignement Secondaire - Salles de TP - Série SMS/ST2S Sciences médicosocial/ scienes et technologies de la santé et du social"
    }
    
    col_SCAP = ["Sous-catégorie d'activité majoritaire", "Sous-catégorie d'activité secondaire"]
    
    OPERAT[col_SCAP] = OPERAT[col_SCAP].replace(SCAP_renaming)

    return OPERAT


############################################################################################################################
#-------------------------------------------------------- SCORE PS --------------------------------------------------------#
############################################################################################################################

def score_PS(OPERAT, CAP_file, s_inf1, s_sup1):

    # Import de la liste des catégories et des bornes
    warnings.simplefilter(action="ignore")
    CAP = pd.read_excel(CAP_file,
                      sheet_name = "Catégories & Sous-catégories",
                      header = 0,
                      skiprows = 0,
                      dtype = object)
    warnings.resetwarnings()
    
    CAP = CAP.map(lambda x: x.replace("’", "'") if isinstance(x, str) else x, na_action="ignore")
    CAP = CAP.map(lambda x: " ".join(x.split()) if isinstance(x, str) else x, na_action="ignore")
    CAP = CAP.map(lambda x: x.strip() if isinstance(x, str) else x, na_action="ignore")
    
    # Liste des CAP et SCAP possibles (v5)
    list_CAP = CAP.loc[CAP["Type"] == "Catégorie"]["Catégories et sous-catégories"].value_counts().index.to_list()
    list_SCAP = CAP.loc[CAP["Type"] == "Sous-Catégorie"]["Catégories et sous-catégories"].value_counts().index.to_list()

    # Nouvelle df pour traiter les données
    df = copy.copy(OPERAT)
    
    # Ajout des codes pour la CAP-M et la CAP-S
    df["Code CAP-M"] = np.nan
    df["Code CAP-M"] = df["Code CAP-M"].astype('object')
    df["Code CAP-S"] = np.nan
    df["Code CAP-S"] = df["Code CAP-S"].astype('object')
    
    list_CAP_M = df["Catégorie d'activité majoritaire"].value_counts().index.to_list()
    list_CAP_S = df["Catégorie d'activité majoritaire"].value_counts().index.to_list()
    list_CAP_OPERAT = list(set(list_CAP_M + list_CAP_S))
    
    # Pour toutes les CAP-M ou CAP-S
    for cap in list_CAP_M or cap in list_CAP_S:
    
        # Si la CAP est une CAP-M et est dans la list des options possibles, la renseigner dans la df
        if cap in list_CAP and cap in list_CAP_M:
            code_cap = CAP.loc[CAP["Type"] == "Catégorie"].loc[CAP["Catégories et sous-catégories"] == cap]["Code"].iloc[0]
            idx = df.loc[df["Catégorie d'activité majoritaire"] == cap].index
            df.loc[idx, "Code CAP-M"] = code_cap
    
        # Idem pour CAP-S
        if cap in list_CAP and cap in list_CAP_S:
            code_cap = CAP.loc[CAP["Type"] == "Catégorie"].loc[CAP["Catégories et sous-catégories"] == cap]["Code"].iloc[0]
            idx = df.loc[df["Catégorie d'activité secondaire"] == cap].index
            df.loc[idx, "Code CAP-S"] = code_cap 

    # Ajout de la colonne P&S
    df["P&S Score"] = 0
    df["P&S Score"] = df["P&S Score"].astype('Int8')
    
    ####################################################
    # PS 1 : Surface totale brute non-nulle renseignées
    ####################################################
    n_test = 1
    idx = df.loc[df["Surface totale brute (m²)"].fillna(0) > 0].index
    df.loc[idx, "P&S Score"] = n_test
    
    ############################################################################
    # PS 2 : CAP-M renseignée et CAP-M dans la liste des options disponbibles ? 
    ############################################################################
    n_test = 2
    df_extract = df.loc[df["P&S Score"] >= n_test-1] # données avec P&S >= 1
    
    list_CAP_M = df_extract["Catégorie d'activité majoritaire"].value_counts().index.to_list()
    CAP_M_no_match = [x for x in list_CAP_M if x not in list_CAP]
    
    # Indices P&S = 2
    idx = df_extract.loc[pd.isnull(df_extract["Catégorie d'activité majoritaire"])].index # CAP-M non renseignées
    idx = idx.union(df_extract.loc[df["Catégorie d'activité majoritaire"].isin(CAP_M_no_match)].index) # CAP-M pas dans la liste
    idx2 = df_extract.index.difference(idx) # Tous les indices pour PS = 1 sauf ceux ci-dessus
    
    # Attribution score
    df.loc[idx2, "P&S Score"] = n_test
    
    #####################################
    # PS 3: les surfaces sont cohérentes
    #####################################
    n_test = 3
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[df_extract["Surface moyenne annuelle"].fillna(0) <= 1.005*df_extract["Surface totale brute (m²)"].fillna(0)].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ################################
    # PS 4: L'EFA n'est pas vacante
    ################################
    n_test = 4
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[df_extract['Surface moyenne annuelle local vacant (m²)'].fillna(0) < 0.99*df_extract["Surface moyenne annuelle"].fillna(0)].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ############################################################
    # PS 5: Période de reporting = 1 an (à 1% i.e. 5 jours près)
    ###########################################################
    n_test = 5
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[(df_extract["Surface moyenne annuelle"] >= 0.99*df_extract["Surface totale brute (m²)"])].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ##################################################################
    # PS 6: Surf brute comprise entre les bornes (1er filtre physique)
    ##################################################################
    n_test = 6
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[df_extract["Surface totale brute (m²)"].between(s_inf1, s_sup1)].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ############################################################
    # PS 7: 2ème filtre physique (bornes dépendante de la CAP-M)
    ############################################################
    n_test = 7
    df_extract = copy.copy(df.loc[df["P&S Score"] >= n_test-1]) # Ici on crée une copie pour éviter les warnings
    
    # Ajout des colonnes pour le calcul de S_inf et S_sup
    df_extract["Surf non vac"] = np.nan # Surface non-vacante EFA
    df_extract["Surf ratio CAP-M/EFA"] = np.nan # Ratio Surf CAP-M/Surf non-vacante EFA
    df_extract["Surf ratio CAP-S/EFA"] = np.nan # Ratio Surf CAP-S/Surf non-vacante EFA
    df_extract["Surf ratio CAP rest/EFA"] = np.nan # Ratio Surf autres CAP/Surf non-vacante EFA
    df_extract["S_inf CAP-M"] = np.nan # Borne inf pour la CAP-M
    df_extract["S_inf CAP-S"] = np.nan # Borne inf pour la CAP-S
    df_extract["S_sup CAP-M"] = np.nan # Borne sup pour la CAP-M
    df_extract["S_sup CAP-S"] = np.nan # Borne sup pour la CAP-S
    
    # On isole les EFA pour lesquelles CAP-M ou S = Local vacant. Elles feront l'objet d'un calcul particulier.
    idx_novac = df_extract.loc[df_extract["Code CAP-M"] != "CAP00"].loc[df_extract["Code CAP-S"] != "CAP00"].index
    
    # Surface EFA non vacante
    df_extract.loc[idx_novac, "Surf non vac"] = (df_extract["Surface moyenne annuelle"].loc[idx_novac].fillna(0) 
                                                 - df_extract["Surface moyenne annuelle local vacant (m²)"].loc[idx_novac].fillna(0))
    
    # Ratio Surf CAP-M/Surf non-vacante EFA
    df_extract.loc[idx_novac, "Surf ratio CAP-M/EFA"] = df_extract["Surface moyenne annuelle catégorie d'activité majoritaire (m²)"].loc[idx_novac].fillna(0)/df_extract["Surf non vac"].loc[idx_novac]
    
    # Ratio Surf CAP-S/Surf non-vacante EFA
    df_extract.loc[idx_novac, "Surf ratio CAP-S/EFA"] = df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_novac].fillna(0)/df_extract["Surf non vac"].loc[idx_novac]
    
    # Ratio Surf autres CAP/Surf non-vacante EFA
    df_extract.loc[idx_novac, "Surf ratio CAP rest/EFA"] = (df_extract["Surf non vac"].loc[idx_novac].fillna(0)
                                                           - df_extract["Surface moyenne annuelle catégorie d'activité majoritaire (m²)"].loc[idx_novac].fillna(0)
                                                           - df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_novac].fillna(0))/df_extract["Surf non vac"].loc[idx_novac]
    
    # On crée 2 dict comprenant la liste des codes CAP pour chaque borne surfacique (inf et sup)
    
    dict_Sinf = {}
    for S in CAP["S_inf (m²)"].value_counts().index.to_list():
        dict_Sinf[S] = CAP.loc[CAP["S_inf (m²)"] == S]["Code"].value_counts().index.to_list()
    # On remplace 'Défaut' par sa valeur (min des autres bornes)
    default_Sinf = min([x for x in CAP["S_inf (m²)"].value_counts().index.to_list() if isinstance(x, int) or isinstance(x, float)])
    dict_Sinf[default_Sinf] = dict_Sinf.pop('Defaut') + dict_Sinf[default_Sinf]
    
    dict_Ssup = {}
    for S in CAP["S_sup (m²)"].value_counts().index.to_list():
        dict_Ssup[S] = CAP.loc[CAP["S_sup (m²)"] == S]["Code"].value_counts().index.to_list()
    # On remplace 'Défaut' par sa valeur (min des autres bornes) et Aucune par la borne sup 1 (définie dans un test P&S) précédent
    default_Ssup = max([x for x in dict_Ssup.keys() if isinstance(x, int) or isinstance(x, float)])
    dict_Ssup[default_Ssup] = dict_Ssup.pop('Defaut') + dict_Ssup[default_Ssup]
    aucune_Ssup = s_sup1
    dict_Ssup[aucune_Ssup] = dict_Ssup.pop('Aucune')
    
    
    # Calcul bornes pour la CAP-M (en excluant les CAP-M ou S = Local vacant)
    for CAP_code in df_extract["Code CAP-M"].loc[idx_novac].value_counts().index.to_list():
        
        # Indices des lignes qui ont cette CAP renseignée en CAP-M (en excluant toujours les CAP-M ou S = Local vacant)
        idx = df_extract.loc[idx_novac].loc[df_extract["Code CAP-M"] == CAP_code].index
        
        # Liste des S_inf et S_sup possibles pour cette CAP
        S_inf_cap = [key for key in dict_Sinf if CAP_code in dict_Sinf[key]][0]
        S_sup_cap = [key for key in dict_Ssup if CAP_code in dict_Ssup[key]][0]
        
        # Renseignement des valeurs dans les colonne ad hoc de df extract
        df_extract.loc[idx , "S_inf CAP-M"] = S_inf_cap
        df_extract.loc[idx , "S_sup CAP-M"] = S_sup_cap
    
    # Calcul bornes pour la CAP-S (idem CAP-M)
    for CAP_code in df_extract["Code CAP-S"].loc[idx_novac].value_counts().index.to_list():
        idx = df_extract.loc[idx_novac].loc[df_extract["Code CAP-S"] == CAP_code].index
        S_inf_cap = [key for key in dict_Sinf if CAP_code in dict_Sinf[key]][0]
        S_sup_cap = [key for key in dict_Ssup if CAP_code in dict_Ssup[key]][0]
        df_extract.loc[idx , "S_inf CAP-S"] = S_inf_cap
        df_extract.loc[idx , "S_sup CAP-S"] = S_sup_cap
    
    # Calcul bornes EFA
    df_extract.loc[idx_novac, "S_inf (m²)"] = (df_extract["Surf ratio CAP-M/EFA"].loc[idx_novac].fillna(0)*df_extract["S_inf CAP-M"].loc[idx_novac].fillna(0)
                                               + df_extract["Surf ratio CAP-S/EFA"].loc[idx_novac].fillna(0)*df_extract["S_inf CAP-S"].loc[idx_novac].fillna(0)
                                              + df_extract["Surf ratio CAP rest/EFA"].loc[idx_novac].fillna(0)*default_Sinf)
    
    df_extract.loc[idx_novac, "S_sup (m²)"] = (df_extract["Surf ratio CAP-M/EFA"].loc[idx_novac].fillna(0)*df_extract["S_sup CAP-M"].loc[idx_novac].fillna(0)
                                               + df_extract["Surf ratio CAP-S/EFA"].loc[idx_novac].fillna(0)*df_extract["S_sup CAP-S"].loc[idx_novac].fillna(0)
                                              + df_extract["Surf ratio CAP rest/EFA"].loc[idx_novac].fillna(0)*default_Ssup)
    
    # Cas particuliers CAP-M ou S = Local vacant
    
    # Cas 1: CAP-M = Local vacant
    
    # Indices des lignes correspondantes
    idx_vac_CAP_M = df_extract.loc[df_extract["Code CAP-M"] == "CAP00"].index
    
    # Surface EFA non vacante
    df_extract.loc[idx_vac_CAP_M, "Surf non vac"] = (df_extract["Surface moyenne annuelle"].loc[idx_vac_CAP_M].fillna(0) 
                                                 - df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_vac_CAP_M].fillna(0))
    
    # Ratio Surf CAP-M/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_M, "Surf ratio CAP-M/EFA"] = 0 # Car CAP-M vacante
    
    # Ratio Surf CAP-S/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_M, "Surf ratio CAP-S/EFA"] = df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_vac_CAP_M].fillna(0)/df_extract["Surf non vac"].loc[idx_vac_CAP_M]
    
    # Ratio Surf autres CAP/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_M, "Surf ratio CAP rest/EFA"] = (df_extract["Surf non vac"].loc[idx_vac_CAP_M].fillna(0)
                                                           - df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_vac_CAP_M].fillna(0))/df_extract["Surf non vac"].loc[idx_vac_CAP_M]
    
    # Calcul bornes pour la CAP-S
    for CAP_code in df_extract["Code CAP-S"].loc[idx_vac_CAP_M].value_counts().index.to_list():
        idx = df_extract.loc[idx_vac_CAP_M].loc[df_extract["Code CAP-S"] == CAP_code].index
        S_inf_cap = [key for key in dict_Sinf if CAP_code in dict_Sinf[key]][0]
        S_sup_cap = [key for key in dict_Ssup if CAP_code in dict_Ssup[key]][0]
        df_extract.loc[idx , "S_inf CAP-S"] = S_inf_cap
        df_extract.loc[idx , "S_sup CAP-S"] = S_sup_cap
    
    # Cas 2: CAP-S = Local vacant
    
    # Indices des lignes correspondantes
    idx_vac_CAP_S = df_extract.loc[df_extract["Code CAP-S"] == "CAP00"].index
    
    # Surface EFA non vacante
    df_extract.loc[idx_vac_CAP_S, "Surf non vac"] = (df_extract["Surface moyenne annuelle"].loc[idx_vac_CAP_S].fillna(0) 
                                                 - df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_vac_CAP_S].fillna(0))
    
    # Ratio Surf CAP-M/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_S, "Surf ratio CAP-M/EFA"] = df_extract["Surface moyenne annuelle catégorie d'activité majoritaire (m²)"].loc[idx_vac_CAP_S].fillna(0)/df_extract["Surf non vac"].loc[idx_vac_CAP_S]
    
    # Ratio Surf CAP-S/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_S, "Surf ratio CAP-S/EFA"] = 0 # Car CAP-S = local vacant
    
    # Ratio Surf autres CAP/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_S, "Surf ratio CAP rest/EFA"] = (df_extract["Surf non vac"].loc[idx_vac_CAP_S].fillna(0)
                                                           - df_extract["Surface moyenne annuelle catégorie d'activité majoritaire (m²)"].loc[idx_vac_CAP_S].fillna(0))/df_extract["Surf non vac"].loc[idx_vac_CAP_S]
    
    # Calcul bornes pour la CAP-M 
    for CAP_code in df_extract["Code CAP-M"].loc[idx_vac_CAP_S].value_counts().index.to_list():
        idx = df_extract.loc[idx_vac_CAP_S].loc[df_extract["Code CAP-M"] == CAP_code].index
        S_inf_cap = [key for key in dict_Sinf if CAP_code in dict_Sinf[key]][0]
        S_sup_cap = [key for key in dict_Ssup if CAP_code in dict_Ssup[key]][0]
        df_extract.loc[idx , "S_inf CAP-M"] = S_inf_cap
        df_extract.loc[idx , "S_sup CAP-M"] = S_sup_cap
    
    # Calcul bornes EFA
    idx_vac = idx_vac_CAP_M.union(idx_vac_CAP_S)
    
    df_extract.loc[idx_vac, "S_inf (m²)"] = (df_extract["Surf ratio CAP-M/EFA"].loc[idx_vac].fillna(0)*df_extract["S_inf CAP-M"].loc[idx_vac].fillna(0)
                                               + df_extract["Surf ratio CAP-S/EFA"].loc[idx_vac].fillna(0)*df_extract["S_inf CAP-S"].loc[idx_vac].fillna(0)
                                              + df_extract["Surf ratio CAP rest/EFA"].loc[idx_vac].fillna(0)*default_Sinf)
    
    df_extract.loc[idx_vac, "S_sup (m²)"] = (df_extract["Surf ratio CAP-M/EFA"].loc[idx_vac].fillna(0)*df_extract["S_sup CAP-M"].loc[idx_vac].fillna(0)
                                               + df_extract["Surf ratio CAP-S/EFA"].loc[idx_vac].fillna(0)*df_extract["S_sup CAP-S"].loc[idx_vac].fillna(0)
                                              + df_extract["Surf ratio CAP rest/EFA"].loc[idx_vac].fillna(0)*default_Ssup)
    
    df_extract.loc[df_extract["S_sup (m²)"].fillna(0) > s_sup1, "S_sup (m²)"] = s_sup1 # Certains bornes sont supérieures à la borne absolues (effets de bord)
    
    # Indices pour lesquels la surf non-vacante est comprise entre les bornes calculées
    idx = df_extract.loc[df_extract["Surf non vac"].between(df_extract["S_inf (m²)"], df_extract["S_sup (m²)"])].index
    
    # On écrit les bornes dans la df principale pour garder une trace (on travaillait sur l'extrait jusqu'ici)
    df["S_inf (m²)"] = np.nan
    df["S_sup (m²)"] = np.nan
    df.loc[df_extract.index, ["S_inf (m²)", "S_sup (m²)"]] = df_extract[["S_inf (m²)", "S_sup (m²)"]]
    
    # On renseigne le score 
    df.loc[idx, "P&S Score"] = n_test
    
    #######################################################################
    # PS 8: Période de reporting commence au 1er janvier (marge de 5 jours)
    #######################################################################
    n_test = 8
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[df_extract['Date début déclaration'].dt.month == 1].loc[df_extract['Date début déclaration'].dt.day <= 5].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ##################################################################
    # Nombres et % de lignes (EFA) qui passent les différents tests PS
    ##################################################################
    PS_scores = df["P&S Score"].value_counts().sort_index()
    PS_cutoff = pd.DataFrame(index = PS_scores.index, columns = ["N pass", "N pass (%)"])
    
    for PS in PS_scores.index:
        n_lines_left = PS_scores.loc[[x for x in PS_scores.index.to_list() if x >= PS]].sum()
        ratio_left = round(n_lines_left/len(df)*100, 2)
        PS_cutoff.loc[PS, ["N pass", "N pass (%)"]] = [n_lines_left, ratio_left]
    
    return df, PS_cutoff
from collections import OrderedDict
import copy
from datetime import datetime
import exrex
import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter
import numpy as np
import os.path
import pandas as pd
from pandastable import Table
from pprint import pprint
import re
import seaborn as sns
import shutil
from termcolor import colored
import time
import tkinter as tk
from tkinter import *
import unicodedata
import uuid
import warnings
import xml.etree.ElementTree as ET

#--------------------------------------------------- PRE CLEANING OF DF ---------------------------------------------------#
"""
DESCRIPTION: This function removes empty rows and rows containing too many NaN values
INPUTS: 
- df: pandas data frame containing the data to correct
OUTPUTS:
- df: corrected data frame 
WORKFLOW:
- Remove extra spaces at the beginning and end of all strings (typos)
- Convert all empty strings to NaN
- Drop (remove) rows that contain only NaN values (empty rows)
"""

def df_pre_cleaning(df, reset_index):
    
    # Suppressoin de tous les espaces en trop au début et à la fin des chaînes de caractères
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x, na_action="ignore")

    # Conversion de toutes les chaînes de caractères vides et des "-" en NaN
    df = df.map(lambda x: np.nan if (x == "" or x == "-") else x, na_action="ignore")
    
    # Suppression des lignes ne contenant que des NaN                  
    df.dropna(axis=0, how="all", inplace=True) 

    if reset_index == True:
        df = df.reset_index(drop=True)
   
    # Return data
    return df


def sym_distrib(mySeries, n_std):

    # Défintion des 4 fonctions
    transforms = {'real': mySeries,
                  'sqrt': np.sqrt(mySeries),
                  'cbrt': np.cbrt(mySeries),
                  'log': np.log(mySeries)}
    
    # Calcul skewness (valeur absolue) pour chacune des 4 distributions obtenues
    skews = {key: abs(val.skew()) for key, val in transforms.items()} 
    
    # Transformation avec la skewness minimum (en valeur abs)
    min_skew = min(skews, key=skews.get) 

    # Distribution symétrique, sa moyenne et son écart-type
    mySeries_sym = transforms[min_skew] # f(Serie)
    mean = mySeries_sym.mean() # Moy de f(Serie)
    std = mySeries_sym.std() # Std de f(Serie)

    # Distribution symétrique filtrée (on retire tout ce qui est au-delà de moyenne +- n*sigma)
    mySeries_sym_filtered = mySeries_sym[mySeries_sym >= mean - n_std*std][mySeries_sym <= mean + n_std*std] # f(Series) filtrée (n_std autour de la moy)
    
    # Calcul des bornes sur la distribution avant transformation
    if not(mySeries_sym_filtered.empty):
        inf = mySeries[mySeries_sym_filtered.idxmin()] # Borne inf à partir du min de f(Series) filtrée
        sup = mySeries[mySeries_sym_filtered.idxmax()] 
    else:
        inf = mySeries.min()
        sup = mySeries.max()

    return inf, sup, mean, std, min_skew



























############################################################################################################################
#------------------------------------------ IMPORT ET CORRECTIONS PRELIMINAIRES -------------------------------------------#
############################################################################################################################

def import_corr_prelim(OPERAT_file, CAP_file):

    # Import données OPERAT
    warnings.simplefilter(action="ignore")
    OPERAT = pd.read_excel(OPERAT_file,
                      sheet_name = 0,
                      header = 0,
                      skiprows = 0,
                      dtype = object)
    warnings.resetwarnings()

    # Lignes vides, Apostrophes, espaces multiples, espaces en début et fin de chaînes de caractères
    OPERAT.dropna(axis=0, how="all", inplace=True) 
    OPERAT = OPERAT.map(lambda x: x.replace("’", "'") if isinstance(x, str) else x, na_action="ignore")
    OPERAT = OPERAT.map(lambda x: " ".join(x.split()) if isinstance(x, str) else x, na_action="ignore")
    OPERAT = OPERAT.map(lambda x: x.strip() if isinstance(x, str) else x, na_action="ignore")


    # Import de la liste des catégories et des bornes
    warnings.simplefilter(action="ignore")
    CAP = pd.read_excel(CAP_file,
                      sheet_name = "Catégories & Sous-catégories",
                      header = 0,
                      skiprows = 0,
                      dtype = object)
    warnings.resetwarnings()
    
    CAP = CAP.map(lambda x: x.replace("’", "'") if isinstance(x, str) else x, na_action="ignore")
    CAP = CAP.map(lambda x: " ".join(x.split()) if isinstance(x, str) else x, na_action="ignore")
    CAP = CAP.map(lambda x: x.strip() if isinstance(x, str) else x, na_action="ignore")
    
    # Liste des CAP et SCAP possibles (v5)
    list_CAP = CAP.loc[CAP["Type"] == "Catégorie"]["Catégories et sous-catégories"].value_counts().index.to_list()
    list_SCAP = CAP.loc[CAP["Type"] == "Sous-Catégorie"]["Catégories et sous-catégories"].value_counts().index.to_list()

    #########################################################################
    # Renomage des CAP pour qu'elles correspondent à la dernière version (v5)
    #########################################################################
    CAP_renaming = {"Enseignement secondaire": "Enseignement Secondaire",
                    "Enseignement supérieur": "Enseignement Supérieur",
                    "Gares routières": "Transport urbain de voyageurs (métro)",
                    "Restauration": "Restauration - Débit de boissons",
                    "Résidences de tourisme et de loisirs": "Résidence de tourisme et loisirs",
                    "Santé et action sociale": "Santé - Etablissements médico-sociaux"
    }
    
    col_CAP = ["Catégorie d'activité majoritaire", "Catégorie d'activité secondaire", "Catégorie d'activité principale"]
    col_SCAP = ["Sous-catégorie d'activité majoritaire", "Sous-catégorie d'activité secondaire"]
    
    OPERAT[col_CAP] = OPERAT[col_CAP].replace(CAP_renaming)
    
    ##########################################################################
    # Renomage des SCAP pour qu'elles correspondent à la dernière version (v5)
    ##########################################################################
    
    # Certaines sous-catégories sont en réalité des catégories. Lorsque c'est le cas, on remplace par la sous-catégorie par défaut.
    
    # On renomme d'abord les SCAP (cette fonction ne va s'appliquer qu'aux SCAP qui sont des CAP)
    OPERAT[col_SCAP] = OPERAT[col_SCAP].replace(CAP_renaming)
    
    # Liste des SCAP renseignées par les assujettis dans OPERAT
    list_SCAP_M = OPERAT["Sous-catégorie d'activité majoritaire"].value_counts().index.to_list()
    list_SCAP_S = OPERAT["Sous-catégorie d'activité secondaire"].value_counts().index.to_list()
    list_SCAP_OPERAT = list(set(list_SCAP_M + list_SCAP_S))
    
    # Pour chaque SCAP renseignée qui est dans la liste des CAP
    for scap in [x for x in list_SCAP_OPERAT if x in list_CAP]:
        
        # On récupère le code de la CAP correspondante
        code_CAP = CAP.loc[CAP["Type"] == "Catégorie"].loc[CAP["Catégories et sous-catégories"] == scap]["Code"].iloc[0]
    
        # Puis on récupère la liste des SCAP correspondant à cette CAP
        scap_list = CAP.loc[CAP["Type"] == "Sous-Catégorie"].loc[CAP["Code"].fillna("").str.contains(code_CAP)]["Catégories et sous-catégories"].to_list()
    
        # On sélectionne la SCAP par défaut (si elle existe, sinon, on passe) et on remplace dans les données OPERAT
        default_scap = [x for x in scap_list if 'defaut' in unicodedata.normalize('NFKD', x).encode('ASCII', 'ignore').decode('ASCII').lower()]
        if len(default_scap) == 1:
            
            default_scap = default_scap[0]
        
            idx_m = OPERAT.loc[OPERAT["Sous-catégorie d'activité majoritaire"] == scap].index
            idx_s = OPERAT.loc[OPERAT["Sous-catégorie d'activité secondaire"] == scap].index
            if len(idx_m) != 0:
                OPERAT.loc[idx_m, "Sous-catégorie d'activité majoritaire"] = default_scap
            if len(idx_s) != 0:
                OPERAT.loc[idx_s, "Sous-catégorie d'activité secondaire"] = default_scap
        
        else:
            continue
    
    # Renomage des SCAP pour qu'elles correspondent à la dernière version (v5)
    SCAP_renaming = {"Activité de santé libérale - Kinésithérapie, Rééducation fonctionnelle": "Activité de santé libérale - Kinésithérapie, Rééducation fonctionnelle,",
                     "Aéroport commercial - Hangars de maintenance aéronautique (gros porteurs) - Densité énergétique Niveau 1 (DE = A W/m²)": "Aéroport commercial - Hangar de maintenance aéronautique (gros porteurs) - Densité énergétique Niveau 1 (DE ≤ A W/m²)",
                     "Aéroport commercial - Hangars de maintenance aéronautique (gros porteurs) - Densité énergétique Niveau 3 (DE = B W/m²)": "Aéroport commercial - Hangars de maintenance aéronautique (gros porteurs) - Densité énergétique Niveau 3 (DE ≥ B W/m²)",
                     "Aéroport commercial - Tri bagages - Densité énergétique Niveau 1 (DE = A W/m²)": "Aéroport commercial - Tri bagages - Densité énergétique Niveau 1 (DE ≤ A W/m²)",
                     "Aéroport commercial - Tri bagages - Densité énergétique Niveau 3 (DE = B W/m²)": "Aéroport commercial - Tri bagages - Densité énergétique Niveau 3 (DE ≥ B W/m²)",
                     "Blanchisserie": "Blanchisserie industrielle (Valeur par défaut)",
                     "Collège": "Enseignement Secondaire (Valeur par défaut)",
                     "Commerces et service de détail - Accessoire de mode (Bijouterie, …) - Zone de vente": "Commerce et service de détail - Accessoire de mode (Bijouterie, …) - Zone de vente",
                     "Commerces et service de détail - Bien être (Sauna et Hammam) - Zone Public": "Commerce et service de détail - Bien être (Sauna -Hammam) - Zone Public",
                     "Commerce et service de détail  - Bien être (Sauna -Hammam) - Zone Public": "Commerce et service de détail - Bien être (Sauna -Hammam) - Zone Public",
                     "Commerces et service de détail - Culture, Média et Loisirs (Libraire, produits culturels, jeux et loisirs...) - Zone de vente": "Commerce et service de détail - Culture, Média et Loisirs (Libraire, produits culturels, jeux et loisirs...) - Zone de vente",
                     "Commerces et service de détail - Equipement de la personne (Vêtements, lingeries, linges de maison, chaussures, maroquinerie et bagages...) - Zone de vente": "Commerce et service de détail - Equipement de la personne (Vêtements, lingeries, linges de maison, chaussures, maroquinerie et bagages...) - Zone de vente",
                     "Commerces et service de détail - Equipement de la personne et Loisirs (Valeur par défaut)": "Commerce et service de détail - Equipement de la personne et Loisirs (Valeur par défaut)",
                     "Commerces et service de détail - Equipement de la personne et Loisirs - Administration et bureaux": "Commerce et service de détail - Equipement de la personne et Loisirs - Administration et bureaux",
                     "Commerces et service de détail - Equipement de la personne et Loisirs - Réserve d'approche": "Commerce et service de détail - Equipement de la personne et Loisirs - Réserve d'approche",
                     "Commerces et service de détail - Numérique et téléphonie - Zone vente": "Commerce et service de détail - Numérique et téléphonie - Zone vente",
                     "Commerces et service de détail - Santé Optique - Zone vente": "Commerce et service de détail - Santé Optique - Zone vente",
                     "Commerces et service de détail - Santé, Soins (Pharmacie, Parapharmacie) - Zone vente": "Commerce et service de détail - Santé, Soins (Pharmacie, Parapharmacie) - Zone vente",
                     "Commerces et service de détail - Service Conseil (Agences de voyages…) - Zone vente": "Commerce et service de détail - Service Conseil (Agences de voyages…) - Zone vente",
                     "Commerces et service de détail - Service Laverie automatique - Zone public": "Commerce et service de détail - Service Laverie automatique - Zone public",
                     "Commerces et service de détail - Service Pressing - Accueil public et process": "Commerce et service de détail - Service Pressing - Accueil public et process",
                     "Commerces et service de détail - Services Equipements de la personne (Cordonnerie, Couturier,…) - Accueil public et process": "Commerce et service de détail - Services Equipements de la personne (Cordonnerie, Couturier,…) - Accueil public et process",
                     "Commerces et service de détail - Soins & Beauté (Beauté & bien être) - Zone soins": "Commerce et service de détail - Soins & Beauté (Beauté & bien être) - Zone soins",
                     "Commerces et service de détail - Soins & Beauté (Parfumerie, cosmétique…) - Zone vente": "Commerce et service de détail - Soins & Beauté (Parfumerie, cosmétique…) - Zone vente",
                     "Commerces et service de détail - Soins de la personne (Coiffeur, Salon d'esthétique, Massage) - Zone vente": "Commerce et service de détail - Soins de la personne (Coiffeur, Salon d'esthétique, Massage) - Zone vente",
                     "Commerces et service de détail - Sports et Outdoor - Zone vente": "Commerce et service de détail - Sports et Outdoor - Zone vente",
                     "Enseignement Secondaire - Lycée d'enseignement général et techhnologique agricole (LGTA) - Lycée d'enseignement professionnel agricole (LEP agricole ) - Toutes séries confondues (Valeur Témoin)": "Enseignement Secondaire - Lycée d'enseignement général et technologique agricole (LGTA) - Lycée d'enseignement professionnel agricole (LEP agricole) - Toutes séries confondues (Valeur Témoin)",
                     "Enseignement Supérieur - Ateliers et Halles techniques avec Process - Niveau 1 (DE = A W/m²)": "Enseignement Supérieur - Ateliers et Halles techniques avec Process - Niveau 1 (DE ≤ A W/m²)",
                     "Enseignement Supérieur - Ateliers et Halles techniques avec Process - Niveau 3 (DE = B W/m²)": "Enseignement Supérieur - Ateliers et Halles techniques avec Process - Niveau 3 (DE ≥ B W/m²)",
                     "Gare ferroviaire - Zones de remisage couvertes et closes": "Gare ferroviaire - Zones de remisage couverte et close",
                     "Lycée d'enseignement général": "Enseignement Secondaire - Lycée d'enseignement général (LG) ou Salles d'enseignement banalisé - Toutes séries confondues (Valeur Témoin)",
                     "Lycée d'enseignement général et technologique – Lycée polyvalent": "Enseignement Secondaire - Lycée d'enseignement général et techhnologique (LGT) - Lycée d'enseignement polyvalent (LEP) - Toutes séries confondues (Valeur Témoin)",
                     "Lycée d'enseignement professionnel": "Enseignement Secondaire - Lycée d'enseignement professionnel",
                     "Salle serveur & Data Center (Valeur par défaut)": "Salle serveur & Data Center - Valeur par défaut",
                     "Stockage à température ambiante": "Logistique température ambiante",
                     "Transport urbain (Valeur par défaut)": "Transport urbain - Valeur par défaut",
                     "nseignement Secondaire - Salles de TP - Série SMS/ST2S Sciences médicosocial/ sciences et technologies de la santé et du social": "Enseignement Secondaire - Salles de TP - Série SMS/ST2S Sciences médicosocial/ scienes et technologies de la santé et du social"
    }
    
    col_SCAP = ["Sous-catégorie d'activité majoritaire", "Sous-catégorie d'activité secondaire"]
    
    OPERAT[col_SCAP] = OPERAT[col_SCAP].replace(SCAP_renaming)

    return OPERAT


############################################################################################################################
#-------------------------------------------------------- SCORE PS --------------------------------------------------------#
############################################################################################################################

def score_PS(OPERAT, CAP_file, s_inf1, s_sup1):

    # Import de la liste des catégories et des bornes
    warnings.simplefilter(action="ignore")
    CAP = pd.read_excel(CAP_file,
                      sheet_name = "Catégories & Sous-catégories",
                      header = 0,
                      skiprows = 0,
                      dtype = object)
    warnings.resetwarnings()
    
    CAP = CAP.map(lambda x: x.replace("’", "'") if isinstance(x, str) else x, na_action="ignore")
    CAP = CAP.map(lambda x: " ".join(x.split()) if isinstance(x, str) else x, na_action="ignore")
    CAP = CAP.map(lambda x: x.strip() if isinstance(x, str) else x, na_action="ignore")
    
    # Liste des CAP et SCAP possibles (v5)
    list_CAP = CAP.loc[CAP["Type"] == "Catégorie"]["Catégories et sous-catégories"].value_counts().index.to_list()
    list_SCAP = CAP.loc[CAP["Type"] == "Sous-Catégorie"]["Catégories et sous-catégories"].value_counts().index.to_list()

    # Nouvelle df pour traiter les données
    df = copy.copy(OPERAT)
    
    # Ajout des codes pour la CAP-M et la CAP-S
    df["Code CAP-M"] = np.nan
    df["Code CAP-M"] = df["Code CAP-M"].astype('object')
    df["Code CAP-S"] = np.nan
    df["Code CAP-S"] = df["Code CAP-S"].astype('object')
    
    list_CAP_M = df["Catégorie d'activité majoritaire"].value_counts().index.to_list()
    list_CAP_S = df["Catégorie d'activité majoritaire"].value_counts().index.to_list()
    list_CAP_OPERAT = list(set(list_CAP_M + list_CAP_S))
    
    # Pour toutes les CAP-M ou CAP-S
    for cap in list_CAP_M or cap in list_CAP_S:
    
        # Si la CAP est une CAP-M et est dans la list des options possibles, la renseigner dans la df
        if cap in list_CAP and cap in list_CAP_M:
            code_cap = CAP.loc[CAP["Type"] == "Catégorie"].loc[CAP["Catégories et sous-catégories"] == cap]["Code"].iloc[0]
            idx = df.loc[df["Catégorie d'activité majoritaire"] == cap].index
            df.loc[idx, "Code CAP-M"] = code_cap
    
        # Idem pour CAP-S
        if cap in list_CAP and cap in list_CAP_S:
            code_cap = CAP.loc[CAP["Type"] == "Catégorie"].loc[CAP["Catégories et sous-catégories"] == cap]["Code"].iloc[0]
            idx = df.loc[df["Catégorie d'activité secondaire"] == cap].index
            df.loc[idx, "Code CAP-S"] = code_cap 

    # Ajout de la colonne P&S
    df["P&S Score"] = 0
    df["P&S Score"] = df["P&S Score"].astype('Int8')
    
    ####################################################
    # PS 1 : Surface totale brute non-nulle renseignées
    ####################################################
    n_test = 1
    idx = df.loc[df["Surface totale brute (m²)"].fillna(0) > 0].index
    df.loc[idx, "P&S Score"] = n_test
    
    ############################################################################
    # PS 2 : CAP-M renseignée et CAP-M dans la liste des options disponbibles ? 
    ############################################################################
    n_test = 2
    df_extract = df.loc[df["P&S Score"] >= n_test-1] # données avec P&S >= 1
    
    list_CAP_M = df_extract["Catégorie d'activité majoritaire"].value_counts().index.to_list()
    CAP_M_no_match = [x for x in list_CAP_M if x not in list_CAP]
    
    # Indices P&S = 2
    idx = df_extract.loc[pd.isnull(df_extract["Catégorie d'activité majoritaire"])].index # CAP-M non renseignées
    idx = idx.union(df_extract.loc[df["Catégorie d'activité majoritaire"].isin(CAP_M_no_match)].index) # CAP-M pas dans la liste
    idx2 = df_extract.index.difference(idx) # Tous les indices pour PS = 1 sauf ceux ci-dessus
    
    # Attribution score
    df.loc[idx2, "P&S Score"] = n_test
    
    #####################################
    # PS 3: les surfaces sont cohérentes
    #####################################
    n_test = 3
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[df_extract["Surface moyenne annuelle"].fillna(0) <= 1.005*df_extract["Surface totale brute (m²)"].fillna(0)].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ################################
    # PS 4: L'EFA n'est pas vacante
    ################################
    n_test = 4
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[df_extract['Surface moyenne annuelle local vacant (m²)'].fillna(0) < 0.99*df_extract["Surface moyenne annuelle"].fillna(0)].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ############################################################
    # PS 5: Période de reporting = 1 an (à 1% i.e. 5 jours près)
    ###########################################################
    n_test = 5
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[(df_extract["Surface moyenne annuelle"] >= 0.99*df_extract["Surface totale brute (m²)"])].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ##################################################################
    # PS 6: Surf brute comprise entre les bornes (1er filtre physique)
    ##################################################################
    n_test = 6
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[df_extract["Surface totale brute (m²)"].between(s_inf1, s_sup1)].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ############################################################
    # PS 7: 2ème filtre physique (bornes dépendante de la CAP-M)
    ############################################################
    n_test = 7
    df_extract = copy.copy(df.loc[df["P&S Score"] >= n_test-1]) # Ici on crée une copie pour éviter les warnings
    
    # Ajout des colonnes pour le calcul de S_inf et S_sup
    df_extract["Surf non vac"] = np.nan # Surface non-vacante EFA
    df_extract["Surf ratio CAP-M/EFA"] = np.nan # Ratio Surf CAP-M/Surf non-vacante EFA
    df_extract["Surf ratio CAP-S/EFA"] = np.nan # Ratio Surf CAP-S/Surf non-vacante EFA
    df_extract["Surf ratio CAP rest/EFA"] = np.nan # Ratio Surf autres CAP/Surf non-vacante EFA
    df_extract["S_inf CAP-M"] = np.nan # Borne inf pour la CAP-M
    df_extract["S_inf CAP-S"] = np.nan # Borne inf pour la CAP-S
    df_extract["S_sup CAP-M"] = np.nan # Borne sup pour la CAP-M
    df_extract["S_sup CAP-S"] = np.nan # Borne sup pour la CAP-S
    
    # On isole les EFA pour lesquelles CAP-M ou S = Local vacant. Elles feront l'objet d'un calcul particulier.
    idx_novac = df_extract.loc[df_extract["Code CAP-M"] != "CAP00"].loc[df_extract["Code CAP-S"] != "CAP00"].index
    
    # Surface EFA non vacante
    df_extract.loc[idx_novac, "Surf non vac"] = (df_extract["Surface moyenne annuelle"].loc[idx_novac].fillna(0) 
                                                 - df_extract["Surface moyenne annuelle local vacant (m²)"].loc[idx_novac].fillna(0))
    
    # Ratio Surf CAP-M/Surf non-vacante EFA
    df_extract.loc[idx_novac, "Surf ratio CAP-M/EFA"] = df_extract["Surface moyenne annuelle catégorie d'activité majoritaire (m²)"].loc[idx_novac].fillna(0)/df_extract["Surf non vac"].loc[idx_novac]
    
    # Ratio Surf CAP-S/Surf non-vacante EFA
    df_extract.loc[idx_novac, "Surf ratio CAP-S/EFA"] = df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_novac].fillna(0)/df_extract["Surf non vac"].loc[idx_novac]
    
    # Ratio Surf autres CAP/Surf non-vacante EFA
    df_extract.loc[idx_novac, "Surf ratio CAP rest/EFA"] = (df_extract["Surf non vac"].loc[idx_novac].fillna(0)
                                                           - df_extract["Surface moyenne annuelle catégorie d'activité majoritaire (m²)"].loc[idx_novac].fillna(0)
                                                           - df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_novac].fillna(0))/df_extract["Surf non vac"].loc[idx_novac]
    
    # On crée 2 dict comprenant la liste des codes CAP pour chaque borne surfacique (inf et sup)
    
    dict_Sinf = {}
    for S in CAP["S_inf (m²)"].value_counts().index.to_list():
        dict_Sinf[S] = CAP.loc[CAP["S_inf (m²)"] == S]["Code"].value_counts().index.to_list()
    # On remplace 'Défaut' par sa valeur (min des autres bornes)
    default_Sinf = min([x for x in CAP["S_inf (m²)"].value_counts().index.to_list() if isinstance(x, int) or isinstance(x, float)])
    dict_Sinf[default_Sinf] = dict_Sinf.pop('Defaut') + dict_Sinf[default_Sinf]
    
    dict_Ssup = {}
    for S in CAP["S_sup (m²)"].value_counts().index.to_list():
        dict_Ssup[S] = CAP.loc[CAP["S_sup (m²)"] == S]["Code"].value_counts().index.to_list()
    # On remplace 'Défaut' par sa valeur (min des autres bornes) et Aucune par la borne sup 1 (définie dans un test P&S) précédent
    default_Ssup = max([x for x in dict_Ssup.keys() if isinstance(x, int) or isinstance(x, float)])
    dict_Ssup[default_Ssup] = dict_Ssup.pop('Defaut') + dict_Ssup[default_Ssup]
    aucune_Ssup = s_sup1
    dict_Ssup[aucune_Ssup] = dict_Ssup.pop('Aucune')
    
    
    # Calcul bornes pour la CAP-M (en excluant les CAP-M ou S = Local vacant)
    for CAP_code in df_extract["Code CAP-M"].loc[idx_novac].value_counts().index.to_list():
        
        # Indices des lignes qui ont cette CAP renseignée en CAP-M (en excluant toujours les CAP-M ou S = Local vacant)
        idx = df_extract.loc[idx_novac].loc[df_extract["Code CAP-M"] == CAP_code].index
        
        # Liste des S_inf et S_sup possibles pour cette CAP
        S_inf_cap = [key for key in dict_Sinf if CAP_code in dict_Sinf[key]][0]
        S_sup_cap = [key for key in dict_Ssup if CAP_code in dict_Ssup[key]][0]
        
        # Renseignement des valeurs dans les colonne ad hoc de df extract
        df_extract.loc[idx , "S_inf CAP-M"] = S_inf_cap
        df_extract.loc[idx , "S_sup CAP-M"] = S_sup_cap
    
    # Calcul bornes pour la CAP-S (idem CAP-M)
    for CAP_code in df_extract["Code CAP-S"].loc[idx_novac].value_counts().index.to_list():
        idx = df_extract.loc[idx_novac].loc[df_extract["Code CAP-S"] == CAP_code].index
        S_inf_cap = [key for key in dict_Sinf if CAP_code in dict_Sinf[key]][0]
        S_sup_cap = [key for key in dict_Ssup if CAP_code in dict_Ssup[key]][0]
        df_extract.loc[idx , "S_inf CAP-S"] = S_inf_cap
        df_extract.loc[idx , "S_sup CAP-S"] = S_sup_cap
    
    # Calcul bornes EFA
    df_extract.loc[idx_novac, "S_inf (m²)"] = (df_extract["Surf ratio CAP-M/EFA"].loc[idx_novac].fillna(0)*df_extract["S_inf CAP-M"].loc[idx_novac].fillna(0)
                                               + df_extract["Surf ratio CAP-S/EFA"].loc[idx_novac].fillna(0)*df_extract["S_inf CAP-S"].loc[idx_novac].fillna(0)
                                              + df_extract["Surf ratio CAP rest/EFA"].loc[idx_novac].fillna(0)*default_Sinf)
    
    df_extract.loc[idx_novac, "S_sup (m²)"] = (df_extract["Surf ratio CAP-M/EFA"].loc[idx_novac].fillna(0)*df_extract["S_sup CAP-M"].loc[idx_novac].fillna(0)
                                               + df_extract["Surf ratio CAP-S/EFA"].loc[idx_novac].fillna(0)*df_extract["S_sup CAP-S"].loc[idx_novac].fillna(0)
                                              + df_extract["Surf ratio CAP rest/EFA"].loc[idx_novac].fillna(0)*default_Ssup)
    
    # Cas particuliers CAP-M ou S = Local vacant
    
    # Cas 1: CAP-M = Local vacant
    
    # Indices des lignes correspondantes
    idx_vac_CAP_M = df_extract.loc[df_extract["Code CAP-M"] == "CAP00"].index
    
    # Surface EFA non vacante
    df_extract.loc[idx_vac_CAP_M, "Surf non vac"] = (df_extract["Surface moyenne annuelle"].loc[idx_vac_CAP_M].fillna(0) 
                                                 - df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_vac_CAP_M].fillna(0))
    
    # Ratio Surf CAP-M/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_M, "Surf ratio CAP-M/EFA"] = 0 # Car CAP-M vacante
    
    # Ratio Surf CAP-S/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_M, "Surf ratio CAP-S/EFA"] = df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_vac_CAP_M].fillna(0)/df_extract["Surf non vac"].loc[idx_vac_CAP_M]
    
    # Ratio Surf autres CAP/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_M, "Surf ratio CAP rest/EFA"] = (df_extract["Surf non vac"].loc[idx_vac_CAP_M].fillna(0)
                                                           - df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_vac_CAP_M].fillna(0))/df_extract["Surf non vac"].loc[idx_vac_CAP_M]
    
    # Calcul bornes pour la CAP-S
    for CAP_code in df_extract["Code CAP-S"].loc[idx_vac_CAP_M].value_counts().index.to_list():
        idx = df_extract.loc[idx_vac_CAP_M].loc[df_extract["Code CAP-S"] == CAP_code].index
        S_inf_cap = [key for key in dict_Sinf if CAP_code in dict_Sinf[key]][0]
        S_sup_cap = [key for key in dict_Ssup if CAP_code in dict_Ssup[key]][0]
        df_extract.loc[idx , "S_inf CAP-S"] = S_inf_cap
        df_extract.loc[idx , "S_sup CAP-S"] = S_sup_cap
    
    # Cas 2: CAP-S = Local vacant
    
    # Indices des lignes correspondantes
    idx_vac_CAP_S = df_extract.loc[df_extract["Code CAP-S"] == "CAP00"].index
    
    # Surface EFA non vacante
    df_extract.loc[idx_vac_CAP_S, "Surf non vac"] = (df_extract["Surface moyenne annuelle"].loc[idx_vac_CAP_S].fillna(0) 
                                                 - df_extract["Surface moyenne annuelle catégorie d'activité secondaire (m²)"].loc[idx_vac_CAP_S].fillna(0))
    
    # Ratio Surf CAP-M/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_S, "Surf ratio CAP-M/EFA"] = df_extract["Surface moyenne annuelle catégorie d'activité majoritaire (m²)"].loc[idx_vac_CAP_S].fillna(0)/df_extract["Surf non vac"].loc[idx_vac_CAP_S]
    
    # Ratio Surf CAP-S/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_S, "Surf ratio CAP-S/EFA"] = 0 # Car CAP-S = local vacant
    
    # Ratio Surf autres CAP/Surf non-vacante EFA
    df_extract.loc[idx_vac_CAP_S, "Surf ratio CAP rest/EFA"] = (df_extract["Surf non vac"].loc[idx_vac_CAP_S].fillna(0)
                                                           - df_extract["Surface moyenne annuelle catégorie d'activité majoritaire (m²)"].loc[idx_vac_CAP_S].fillna(0))/df_extract["Surf non vac"].loc[idx_vac_CAP_S]
    
    # Calcul bornes pour la CAP-M 
    for CAP_code in df_extract["Code CAP-M"].loc[idx_vac_CAP_S].value_counts().index.to_list():
        idx = df_extract.loc[idx_vac_CAP_S].loc[df_extract["Code CAP-M"] == CAP_code].index
        S_inf_cap = [key for key in dict_Sinf if CAP_code in dict_Sinf[key]][0]
        S_sup_cap = [key for key in dict_Ssup if CAP_code in dict_Ssup[key]][0]
        df_extract.loc[idx , "S_inf CAP-M"] = S_inf_cap
        df_extract.loc[idx , "S_sup CAP-M"] = S_sup_cap
    
    # Calcul bornes EFA
    idx_vac = idx_vac_CAP_M.union(idx_vac_CAP_S)
    
    df_extract.loc[idx_vac, "S_inf (m²)"] = (df_extract["Surf ratio CAP-M/EFA"].loc[idx_vac].fillna(0)*df_extract["S_inf CAP-M"].loc[idx_vac].fillna(0)
                                               + df_extract["Surf ratio CAP-S/EFA"].loc[idx_vac].fillna(0)*df_extract["S_inf CAP-S"].loc[idx_vac].fillna(0)
                                              + df_extract["Surf ratio CAP rest/EFA"].loc[idx_vac].fillna(0)*default_Sinf)
    
    df_extract.loc[idx_vac, "S_sup (m²)"] = (df_extract["Surf ratio CAP-M/EFA"].loc[idx_vac].fillna(0)*df_extract["S_sup CAP-M"].loc[idx_vac].fillna(0)
                                               + df_extract["Surf ratio CAP-S/EFA"].loc[idx_vac].fillna(0)*df_extract["S_sup CAP-S"].loc[idx_vac].fillna(0)
                                              + df_extract["Surf ratio CAP rest/EFA"].loc[idx_vac].fillna(0)*default_Ssup)
    
    df_extract.loc[df_extract["S_sup (m²)"].fillna(0) > s_sup1, "S_sup (m²)"] = s_sup1 # Certains bornes sont supérieures à la borne absolues (effets de bord)
    
    # Indices pour lesquels la surf non-vacante est comprise entre les bornes calculées
    idx = df_extract.loc[df_extract["Surf non vac"].between(df_extract["S_inf (m²)"], df_extract["S_sup (m²)"])].index
    
    # On écrit les bornes dans la df principale pour garder une trace (on travaillait sur l'extrait jusqu'ici)
    df["S_inf (m²)"] = np.nan
    df["S_sup (m²)"] = np.nan
    df.loc[df_extract.index, ["S_inf (m²)", "S_sup (m²)"]] = df_extract[["S_inf (m²)", "S_sup (m²)"]]
    
    # On renseigne le score 
    df.loc[idx, "P&S Score"] = n_test
    
    #######################################################################
    # PS 8: Période de reporting commence au 1er janvier (marge de 5 jours)
    #######################################################################
    n_test = 8
    df_extract = df.loc[df["P&S Score"] >= n_test-1]
    
    # Indices
    idx = df_extract.loc[df_extract['Date début déclaration'].dt.month == 1].loc[df_extract['Date début déclaration'].dt.day <= 5].index
    
    # Attribution score
    df.loc[idx, "P&S Score"] = n_test
    
    ##################################################################
    # Nombres et % de lignes (EFA) qui passent les différents tests PS
    ##################################################################
    PS_scores = df["P&S Score"].value_counts().sort_index()
    PS_cutoff = pd.DataFrame(index = PS_scores.index, columns = ["N pass", "N pass (%)"])
    
    for PS in PS_scores.index:
        n_lines_left = PS_scores.loc[[x for x in PS_scores.index.to_list() if x >= PS]].sum()
        ratio_left = round(n_lines_left/len(df)*100, 2)
        PS_cutoff.loc[PS, ["N pass", "N pass (%)"]] = [n_lines_left, ratio_left]
    
    return df, PS_cutoff
