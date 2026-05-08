#################################################
# Générateur de tâches - v0.2
#################################################

import pandas as pd
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font
import numpy as np
import pulp as pl
import yaml
import os

print()
print("-----------------------------------------------------------------------------------------------")
print()
print("Générateur de tâches - v0.2")
print()
print("-----------------------------------------------------------------------------------------------")
print()

#################################################
# Chargement des paramètres depuis le fichier YAML
#################################################

#-----------------------------------------------
# Gestion des fichiers
#-----------------------------------------------

# Chargement du fichier de configuration
with open('config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.safe_load(f)

# Fichiers d'entrée/sortie
input_file = config['files']['input']
output_file = config['files']['output']

#-----------------------------------------------
# Facteurs à inclure dans la fonction objectif
#-----------------------------------------------

# Résoudre le problème exact
probleme_exact = config['optimization']['probleme_exact']

# Prendre en compte les libérations si probleme_exact n'est pas activé
if probleme_exact:
    liberation_enabled = False # Ne peut pas être activé si le problème exact est activé
else:
    liberation_enabled = config['optimization']['liberation_enabled']

# Optimisation lexicographique
lexicographic_optimization = config['optimization']['lexicographic_optimization']
min_lexico = config['optimization']['min_lexico']

# Prendre en compte le minimum plusieurs fois dans la fonction objectif (inutile si lexicographic_optimization est activé)
min_prof_enabled = config['optimization']['min_prof_enabled']
if min_prof_enabled:
    nbr_min_count = config['optimization']['nbr_min_count']

# Autoriser les préférences négatives ou les compenser
allow_negative_preferences = config['optimization']['allow_negative_preferences']
enforce_positive_with_negative = config['optimization']['enforce_positive_with_negative']

# Nombre maximum d'heures de cours par prof
max_hours = config['optimization']['max_hours']

#-----------------------------------------------
# Paramètres pour la génération de tâches alternatives
#-----------------------------------------------

# Nbr de tâches supplémentaires à générer
nbr_taches_alternatives = config['taches_alternatives']['nbr_taches_alternatives']
# Facteur de diversité
diversite_factor = config['taches_alternatives']['diversite']

#-----------------------------------------------
# Paramètres du solveur
#-----------------------------------------------

# Limite de temps pour le solveur en secondes
solver_time_limit = config['solver']['time_limit']

# Tolérance pour le solveur
solver_gap = config['solver']['gap']

# Afficher le log dans le terminal
print_solver_log = config['solver']['afficher_log']

# Sauvegarder le log dans un fichier
save_solver_log = config['solver']['sauvegarder_log']
if save_solver_log:
    solver_log_file = config['solver']['fichier_log']
else:
    solver_log_file = None

#################################################
# Chargement des données depuis le fichier Excel
#################################################

# Fonction de chargement des données depuis le fichier Excel
def load_data(tache_file):
    # Charge les feuilles du fichier Excel
    dfs = pd.read_excel(tache_file, sheet_name=["PROF","COURS","PREF","MAX_NB_GR","ATTRIB_PREALABLE"], header=None)

    df_prof = dfs["PROF"]
    df_cours = dfs["COURS"]
    df_pref = dfs["PREF"]
    df_pref = df_pref.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int) # Mettre float si on veut accepter les préférences réelles
    df_max_nb_gr = dfs["MAX_NB_GR"]
    df_max_nb_gr = df_max_nb_gr.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    df_attrib_preal = dfs["ATTRIB_PREALABLE"]
    df_attrib_preal = df_attrib_preal.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

    # Extraction des données de la page PROF
    nbr_prof = int(df_prof.iat[0, 0])
    
    # Extraire toutes les colonnes en une fois
    list_prof = df_prof.iloc[1:nbr_prof+1, 0].to_numpy(dtype=str)
    liberation_prof = df_prof.iloc[1:nbr_prof+1, 1].to_numpy(dtype=float)
    ci_cible = df_prof.iloc[1:nbr_prof+1, 2].to_numpy(dtype=float) # Non utilisé (on utilise la colonne directement dans ci_min et ci_max)
    
    # Calcul de ci_min et ci_max
    ci_min = (df_prof.iloc[1:nbr_prof+1, 2] - df_prof.iloc[1:nbr_prof+1, 3]).to_numpy(dtype=float)
    ci_max = (df_prof.iloc[1:nbr_prof+1, 2] + df_prof.iloc[1:nbr_prof+1, 4]).to_numpy(dtype=float)
    
    prep_min = df_prof.iloc[1:nbr_prof+1, 5].to_numpy(dtype=int) # Non utilisé
    prep_max = df_prof.iloc[1:nbr_prof+1, 6].to_numpy(dtype=int)
    gr_max = df_prof.iloc[1:nbr_prof+1, 7].to_numpy(dtype=int)
    ci_ant = df_prof.iloc[1:nbr_prof+1, 10].to_numpy(dtype=float)

    # Extraction des données de la page COURS
    nbr_cours = int(df_cours.iat[0, 0])
    
    list_cours = df_cours.iloc[1:nbr_cours+1, 0].to_numpy(dtype=str)
    periodes_cours = df_cours.iloc[1:nbr_cours+1, 1].to_numpy(dtype=int)
    groupes_cours = df_cours.iloc[1:nbr_cours+1, 2].to_numpy(dtype=int)
    etudiants_cours = df_cours.iloc[1:nbr_cours+1, 3].to_numpy(dtype=int)

    # Extraction des préférences, nombres max de groupes et attributions préalables
    pref_prof = df_pref.iloc[1:nbr_prof+1, 1:nbr_cours+1].to_numpy()
    max_nb_gr = df_max_nb_gr.iloc[1:nbr_prof+1, 1:nbr_cours+1].to_numpy()
    attrib_preal = df_attrib_preal.iloc[1:nbr_prof+1, 1:nbr_cours+1].to_numpy()

    return {
        'professors': {
            'list': list_prof,
            'nbr': nbr_prof,
            'liberation': liberation_prof,
            'ci_cible': ci_cible,
            'ci_min': ci_min,
            'ci_max': ci_max,
            'prep_min': prep_min,
            'prep_max': prep_max,
            'gr_max': gr_max,
            'ci_ant': ci_ant
        },
        'courses': {
            'list': list_cours,
            'nbr': nbr_cours,
            'periodes': periodes_cours,
            'groupes': groupes_cours,
            'etudiants': etudiants_cours
        },
        'preferences': {
            'prof': pref_prof,
            'max_nb_gr': max_nb_gr,
            'attrib_preal': attrib_preal
        }
    }

# Chargement des données depuis le fichier Excel
data = load_data(input_file)

# Extraction des données
nbr_prof = data['professors']['nbr']
list_prof = data['professors']['list']
liberation_prof = data['professors']['liberation']
ci_cible = data['professors']['ci_cible']
ci_min = data['professors']['ci_min']
ci_max = data['professors']['ci_max']
prep_min = data['professors']['prep_min']
prep_max = data['professors']['prep_max']
gr_max = data['professors']['gr_max']
ci_ant = data['professors']['ci_ant']

nbr_cours = data['courses']['nbr']
list_cours = data['courses']['list']
periodes_cours = data['courses']['periodes']
groupes_cours = data['courses']['groupes']
etudiants_cours = data['courses']['etudiants']

pref_prof = data['preferences']['prof']
max_nb_gr = data['preferences']['max_nb_gr']
attrib_preal = data['preferences']['attrib_preal']

# Liste des indices (prof,cours,nbr de groupes) faisant l'objet d'attributions préalables
attrib_preal_indices = [(i, j, attrib_preal[i][j]) for i in range(nbr_prof) for j in range(nbr_cours) if attrib_preal[i][j] != 0]

#################################################
# Définition du problème de programmation linéaire
#################################################

#-----------------------------------------------
# Fonction de calcul des pénalités au carré
#-----------------------------------------------

# Pré-calcul des préférences pénalisées au carré pour tous les (i,j). Évite les recalculs inutiles
pref_matrix = np.array(1 - (2 - pref_prof)**2/12 - (2 - pref_prof)/6, dtype=float)

#-----------------------------------------------
# Création du problème et des variables de décision
#-----------------------------------------------

# Création du problème
prob = pl.LpProblem("probleme_tache", pl.LpMaximize)

# Variables de décision y[i][j]: le cours j est-il attribué au prof i (1 si oui, 0 sinon)
y = pl.LpVariable.dicts("y", ((i,j) for i in range(nbr_prof) for j in range(nbr_cours)), cat='Binary')

# Variables de décision x[i][j]: nombre de groupes du cours j attribués au prof i
x = pl.LpVariable.dicts("x", ((i,j) for i in range(nbr_prof) for j in range(nbr_cours)), lowBound=0, cat='Integer') # upBound définies plus tard

#-----------------------------------------------
# Construction de la fonction objectif (selon les paramètres du fichier de configuration)
#-----------------------------------------------

def base_obj_term(i, j):
    term = x[(i,j)] * periodes_cours[j] * pref_matrix[i, j] # Fonction de base à optimiser : nbr de groupes du cours j * durée cours j * préférence pour le cours j
    if liberation_enabled:
        if liberation_prof[i] != 0: # Éviter la division par zéro
            term = term / liberation_prof[i] # Prendre en compte la libération du prof
    return term

# Pré-calcul de tous les termes de la fonction objectif - Évite les recalculs inutiles dans chaque contrainte
obj_terms = {(i, j): base_obj_term(i, j) for i in range(nbr_prof) for j in range(nbr_cours)}

# Définition du problème exact, où la moyenne pondérée est maximisée.
# Attention, aucun prof ne peut avoir 0 heure de cours, sinon la moyenne pondérée n'est pas définie
if probleme_exact:
    avg = pl.LpVariable.dicts("avg", range(nbr_prof), lowBound=-1, upBound=1, cat='Continuous') # Moyennes pondérées des préférences des prof

    # On encode le nombre d'heures de cours totales d'un prof en binaire. Permet de réduire le nombre de variables
    K = int(max_hours).bit_length() # Nombre de bits nécessaires pour représenter max_hours
    b = pl.LpVariable.dicts("b", ((i,k) for i in range(nbr_prof) for k in range(K)), cat='Binary') # b[(i,k)] est la k-ième bit du nombre d'heures de cours totales du prof i
    v = pl.LpVariable.dicts("v", ((i,k) for i in range(nbr_prof) for k in range(K)), lowBound=-1, upBound=1, cat='Continuous') # v[(i,k)] = avg[i] * b[(i,k)]

    # On linéarise la moyenne pondérée
    for i in range(nbr_prof):
        prob += pl.lpSum((1 << k) * b[(i,k)] for k in range(K)) == pl.lpSum(periodes_cours[j] * x[(i,j)] for j in range(nbr_cours)) # Le nombre d'heures (converti du binaire) attribuées au prof i doit être égal à la somme des périodes des cours qui lui sont attribués
        prob += pl.lpSum(obj_terms[(i,j)] for j in range(nbr_cours)) == pl.lpSum((1 << k) * v[(i,k)] for k in range(K))

        for k in range(K):
            prob += v[(i,k)] <= b[(i,k)]
            prob += v[(i,k)] >= -b[(i,k)]
            prob += v[(i,k)] <= avg[i] + (1 - b[(i,k)])
            prob += v[(i,k)] >= avg[i] - (1 - b[(i,k)])

    # Fonction objectif
    objective_fct = pl.lpSum(avg[i] for i in range(nbr_prof)) # Objectif: maximiser la moyenne des préférences pondérées

    # Si min_prof_enabled ou lexicographic_optimization est activé, on ajoute une variable pour le minimum des préférences
    if min_prof_enabled or lexicographic_optimization:
        min_pref_value = pl.LpVariable("min_pref_value", lowBound=-1,upBound=1, cat='Continuous') # La préférence pondérée minimale parmi les profs
        
        # Faire en sorte que min_pref_value soit le minimum des avg[i]
        for i in range(nbr_prof):
            prob += min_pref_value <= avg[i]

        # Si min_prof_enabled, la fonction objectif inclut le minimum un certain nombre de fois et la moyenne des préférences pondérées.
        if min_prof_enabled and not lexicographic_optimization:
            objective_fct += nbr_min_count * min_pref_value

# Sinon, définition du problème approché.
# On maximise la somme des préférences pondérées, sans faire de moyenne sur le nombre d'heures enseignées
# Approximation raisonnable et généralement beaucoup plus rapide que le problème exact
else:
    objective_fct = pl.lpSum(obj_terms.values()) # Objectif: maximiser la somme des préférences, non pondérée

    # Si min_prof_enabled ou lexicographic_optimization est activé, on ajoute une variable pour le minimum des préférences
    if min_prof_enabled or lexicographic_optimization:
        if liberation_enabled: # Il faut, par précaution, des bornes plus larges pour min_pref_value si la libération est prise en compte
            min_pref_value = pl.LpVariable("min_pref_value", lowBound= -1.5 * max_hours, upBound= 1.5 * max_hours, cat='Continuous') # La préférence minimale parmi les profs. Estimation très approximative des bornes. Le pire cas ne devrait pas être plus que 1.5 * max_hours en considérant les libérations
        else:
            min_pref_value = pl.LpVariable("min_pref_value", lowBound= - max_hours, upBound= max_hours, cat='Continuous') # La préférence minimale parmi les profs

        # Faire en sorte que min_pref_value soit le minimum
        for i in range(nbr_prof):
            prob += min_pref_value <= pl.lpSum(obj_terms[(i, j)] for j in range(nbr_cours))

        # Si min_prof_enabled, la fonction objectif inclut le minimum plusieurs fois ainsi que la somme totale des termes.
        if min_prof_enabled and not lexicographic_optimization:
            objective_fct += nbr_min_count * min_pref_value

# Ajout de la fonction objectif au problème
prob += objective_fct

#-----------------------------------------------
# Contraintes d'optimisation de base
#-----------------------------------------------

# Contrainte d'optimisation: respecter le nombre maximum de groupes par prof pour chacun des cours
# Si max_nb_gr[i][j] = 0, alors y[i][j] et x[i][j] valent 0 (pas de groupe attribué). Sinon, y[i][j] peut être 1 au maximum et x[i][j] peut être max_nb_gr[i][j] au maximum.
for i in range(nbr_prof):
    for j in range(nbr_cours):
        x[(i,j)].upBound = min(max_nb_gr[i][j], groupes_cours[j])
        y[(i,j)].upBound = int(max_nb_gr[i][j] > 0)

# Si on n'autorise pas les préférences négatives, alors x[i][j] et y[i][j] doivent être 0 pour les cours à préférence négative
if not allow_negative_preferences:
    for i in range(nbr_prof):
        for j in range(nbr_cours):
            if pref_prof[i, j] < 0:
                x[(i,j)].upBound = 0
                y[(i,j)].upBound = 0

# Si on autorise les préférences négatives mais qu'on veut s'assurer qu'elles soient compensées par des préférences positives
if allow_negative_preferences and enforce_positive_with_negative:
    for i in range(nbr_prof):
        prob += pl.lpSum(y[(i,j)] * pref_prof[i][j] for j in range(nbr_cours)) >= 0

# On lie x et y
# On s'assure que le nombre de groupes attribués ne dépasse pas max_nb_gr[i][j]
# On omet les paires (i,j) avec upBound=0 car elles sont déjà fixées à 0 et n'ont pas besoin de contraintes supplémentaires
for i in range(nbr_prof):
    for j in range(nbr_cours):
        if x[(i,j)].upBound > 0:
            prob += x[(i,j)] >= y[(i,j)]
            prob += x[(i,j)] <= y[(i,j)] * max_nb_gr[i][j]

# Tous les groupes de chaque cours doivent être attribués
for j in range(nbr_cours):
    prob += pl.lpSum(x[(i,j)] for i in range(nbr_prof)) == groupes_cours[j]

#-----------------------------------------------
# Contraintes liées à la CI
#-----------------------------------------------

# Fonctions intervenants dans le calcul de la CI
# hc[i] = heures de cours pour le prof i
hc = {i: pl.lpSum(periodes_cours[j] * x[(i,j)] for j in range(nbr_cours)) for i in range(nbr_prof)}

# hp[i] = heures de préparation pour le prof i
hp = {i: pl.lpSum(periodes_cours[j] * y[(i,j)] for j in range(nbr_cours)) for i in range(nbr_prof)}

# nes[i] = nombre d'étudiants pour le prof i
nes = {i: pl.lpSum(etudiants_cours[j] * x[(i,j)] for j in range(nbr_cours)) for i in range(nbr_prof)}

# pes[i] = périodes × étudiants pour le prof i
pes = {i: pl.lpSum(periodes_cours[j] * x[(i,j)] * etudiants_cours[j] for j in range(nbr_cours)) for i in range(nbr_prof)}

# nbr_prep[i] = nombre de préparations pour le prof i
nbr_prep = {i: pl.lpSum(y[(i,j)] for j in range(nbr_cours)) for i in range(nbr_prof)}

# Le calcul de la CI contient des contraintes conditionnelles qu'il faut linéariser.
# On utilise la méthode du Big M

# Linéarisation de la condition nes >= 75
M1 = 160 # >= à la valeur maximale possible de nes(i). On impose nes <= 160 pour éviter les complications de linéarisation au-delà de 160. Un tel scénario serait probablement rejeté dans tous les cas

# Variables binaires z[i]: z[i]=1 si nes(i) >= 75, 0 sinon
z = pl.LpVariable.dicts("z", range(nbr_prof), cat='Binary')

# Variables entières w[i]: w[i]=nes(i) si nes(i) >= 75, 0 sinon
w = pl.LpVariable.dicts("w", range(nbr_prof), lowBound=0, upBound=M1, cat='Continuous') # En théorie, la cat de w[i] devrait être entier, mais sa définition implique déjà qu'il soit entier. Les variables continues sont plus faciles à résoudre.

# Contraintes pour w[i] = nes[i] si nes[i] >= 75, sinon w[i] = 0 avec la méthode Big M
for i in range(nbr_prof):
    # Bornes pour nes[i] pour restreindre le domaine de recherche
    prob += nes[i] <= 160 # On impose que nes[i] ne dépasse pas 160 pour éviter les complications de linéarisation au-delà de 160.

    # Si le professeur a des attributions préalables, on peut calculer une borne inférieure pour nes[i]
    min_nes = sum(etudiants_cours[j] * attrib_preal[i][j] for j in range(nbr_cours) if attrib_preal[i][j] > 0)
    if min_nes >= 75:
        z[i].lowBound = 1 # nes[i] >= 75 assuré, donc z = 1

    # z[i]=1 si nes[i] >= 75, sinon z[i]=0
    prob += nes[i] <= 74 + M1 * z[i]  # Si nes[i] >= 75, alors z[i]=1
    prob += nes[i] >= 75 - 75 * (1 - z[i]) # Si nes[i] < 75, alors z[i]=0

    # w[i] = nes[i] si z[i]=1, sinon w[i]=0
    prob += w[i] <= M1 * z[i] # Force w[i]=0 si z[i]=0
    prob += w[i] >= nes[i] - M1 * (1 - z[i]) # Force w[i]=nes[i] si z[i]=1
    prob += w[i] <= nes[i]

# Linéarisation de la condition pes >= 416
# M2 >= la valeur maximale possible de pes(i) - 415
# Calculé par professeur, à partir de la formule de la ci: ci <= ci_max, hc et hp >= nbr_heures_min et nes >= 0
nbr_heures_min = min(periodes_cours) # Cours avec le moins de périodes, pour la majoration de pes 
M2 = [ max(int(np.floor((ci_max[i] - 2.1 * nbr_heures_min - 415 * 0.04) / 0.07)) + 1, 0) for i in range(nbr_prof) ] # +1 par sécurité

for i in range(len(M2)):
    if M2[i] > 0:
        M2[i] = min(M2[i], 385) # Remplacer la borne par 385 (= 800 - 415) si la borne sup précédemment calculée est trop grande (car par exemple trop grande ci_max). Calculé sur la base de 20h x 40 étudiants.

# Variables binaires t[i]: t[i]=1 si pes(i) >= 416, 0 sinon
t = pl.LpVariable.dicts("t", range(nbr_prof), cat='Binary')

# Variables entières m[i]: m[i]=pes(i) - 415 si pes(i) >= 416, 0 sinon
m = pl.LpVariable.dicts("m", range(nbr_prof), lowBound=0, cat='Continuous') # En théorie, la cat de m[i] devrait être entier, mais sa définition implique déjà qu'il soit entier. Les variables continues sont plus faciles à résoudre.

# Contraintes pour m[i] = pes[i] - 415 si pes[i] >= 416, sinon m[i] = 0
for i in range(nbr_prof):
    m[i].upBound = M2[i]
    prob += m[i] >= pes[i] - 415

    if M2[i] > 0: # Si le pes est peut-être >= 416
        prob += pes[i] >= 416 * t[i]
        prob += pes[i] <= 415 + M2[i] * t[i]
        prob += m[i] <= pes[i] - 415 * t[i]
        prob += m[i] <= M2[i] * t[i]
    else: # Sinon, pes[i] ne peut pas dépasser 415, donc m[i] doit être 0
        t[i].upBound = 0
        # m[i] est automatiquement fixé à 0 par sa upBound.

# Calcul de la CI pour chaque prof
ci = {i: hc[i] * 1.2 + hp[i] * 0.9 + pes[i] * 0.04 + m[i] * 0.03 + w[i] * 0.01 for i in range(nbr_prof)}
# Ce n'est pas la formule complète de la CI: il manque la contribution des nes >= 160. On suppose que ça n'arrive pas pour éviter la pénalité au carré qu'il faudrait linéariser
# Le calcul final de la CI dans le fichier Excel prend en compte la pénalité au carré et est donc correct quoi qu'il arrive

# CI des profs entre min et max
for i in range(nbr_prof):
    prob += ci[i] >= ci_min[i]
    prob += ci[i] <= ci_max[i]

# Nbr d'heures entre 0 et max_hours (défini dans config.yaml)
for i in range(nbr_prof):
    prob += hc[i] <= max_hours

# Nbr total de groupe max par prof entre 0 et gr_max
for i in range(nbr_prof):
    prob += pl.lpSum(x[(i,j)] for j in range(nbr_cours)) <= gr_max[i]

# Nbr de préparations max par prof
for i in range(nbr_prof):
    prob += nbr_prep[i] <= prep_max[i]

# Attribution préalables (à partir de la liste des indices attrib_preal_indices)
for i, j, nb_groupes in attrib_preal_indices:
    y[(i,j)].lowBound = 1
    y[(i,j)].upBound = 1
    x[(i,j)].lowBound = nb_groupes
    # x[(i,j)].upBound est déjà défini par max_nb_gr[i][j] ou groupes_cours[j].
    # Attention: s'assurer que le max_nb_gr[i][j] est rempli pour les attributions préalables

#################################################
# Traitement des résultats
# On arrondit les valeurs qui devraient être entières pour éviter les potentielles approximations du solveur
# Assure un résultat valide dans le fichier Excel
#################################################

def get_val(v): # Retourne v arrondi à l'entier, ou 0 si None - Au cas où la valeur retournée par le solver serait None
    return 0 if v.varValue is None else round(v.varValue)

def round_xy(): # Stocke dans un np.array les valeurs arrondies des variables de décision
    global y_output, x_output
    y_output = np.array([[get_val(y[(i,j)]) for j in range(nbr_cours)] for i in range(nbr_prof)])
    x_output = np.array([[get_val(x[(i,j)]) for j in range(nbr_cours)] for i in range(nbr_prof)])

# Calcul des composantes de la CI à partir des variables de décision arrondies

def hc_output(i):
    # sum_j de periodes_cours[j] * x_output[i,j]
    return int(np.dot(periodes_cours, x_output[i]))

def hp_output(i):
    # sum_j de periodes_cours[j] * y_output[i,j]
    return int(np.dot(periodes_cours, y_output[i]))

def nes_output(i):
    # sum_j de etudiants_cours[j] * x_output[i,j]
    return int(np.dot(etudiants_cours, x_output[i]))

coeff = periodes_cours * etudiants_cours # Pré-calcul des coefficients pour pes_output
def pes_output(i):
    # sum_j de periodes[j] * etudiants[j] * x_output[i,j]
    return int(np.dot(coeff, x_output[i]))

def nbr_prep_output(i):
    # sum_j de y_output[i,j]
    return int(np.sum(y_output[i]))

def ci_output(i):
    if pes_output(i) >= 416:
        total = hc_output(i) * 1.2 + hp_output(i) * 0.9 + 415 * 0.04 + (pes_output(i)-415) * 0.07
    else:
        total = hc_output(i) * 1.2 + hp_output(i) * 0.9 + pes_output(i) * 0.04
    if nes_output(i) >= 75:
        total += nes_output(i) * 0.01
    if nes_output(i) >= 161:
        total += (nes_output(i) - 160)**2 * 0.1
    return total

#################################################
# Exportation des résultats dans un fichier Excel
#################################################

# Fonction d'exportation des résultats dans un fichier Excel
def save_excel_file(writer, numero_tache): # Exportation des résultats dans un fichier Excel comme nouvelle feuille
    round_xy() # Arrondir les valeurs des variables de décision

    # Création du DataFrame des résultats
    df_tache = pd.DataFrame()

    # Remplissage du DataFrame avec les résultats
    # Nom des profs
    df_tache['Prof'] = list_prof

    # Nombre de groupes attribués par cours
    for j in range(nbr_cours):
        col_name = list_cours[j]
        df_tache[col_name] = [x_output[i,j] for i in range(nbr_prof)]

    # Libération
    df_tache['% cours'] = liberation_prof 

    # CI des profs
    ci_values = []
    ci_pleine_values = []
    ci_annee_values = []

    # Pénalités
    pref_pond_heure_values = []

    # Autres indicateurs
    nbr_prep_values = []
    nbr_gr_values = []
    nbr_heures_values = []

    # On ajoute les valeurs aux différentes listes
    for i in range(nbr_prof):
        ci_values.append(ci_output(i))
        ci_pleine_values.append(ci_output(i) + 40 * (1-liberation_prof[i]))
        ci_annee_values.append(ci_ant[i] + ci_pleine_values[i])
        nbr_prep_values.append(nbr_prep_output(i))
        nbr_gr_values.append(sum(x_output[i,j] for j in range(nbr_cours)))
        nbr_heures_values.append(hc_output(i))
        if nbr_heures_values[i] == 0:
            pref_pond_heure_values.append(0)
        else:
            pref_pond_heure_values.append(round(sum(periodes_cours[j] * x_output[i,j] * pref_matrix[i, j]/nbr_heures_values[i] for j in range(nbr_cours)),2))

    # Ajout des colonnes de CI, pénalités et autres indicateurs au DataFrame
    df_tache['CI Cours'] = ci_values
    df_tache['CI Pleine'] = ci_pleine_values
    df_tache['CI Année'] = ci_annee_values
    df_tache['Nbr. Prép.'] = nbr_prep_values
    df_tache['Nbr. Groupes'] = nbr_gr_values
    df_tache['Nbr. Périodes'] = nbr_heures_values
    df_tache['Préf.'] = pref_pond_heure_values

    # Nom de la feuille Excel. "1" pour la première, "2", "3", etc. pour les alternatives
    sheet_name = str(numero_tache + 1)

    # Mise en forme et enregistrement du fichier Excel
    df_tache.to_excel(writer, sheet_name=sheet_name, index=False)
    sheet = writer.sheets[sheet_name]

    # Coloration des cellules contenant des préférences négatives
    couleur_m1 = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid") # Gris clair pour les préférences de -1
    couleur_m2 = PatternFill(start_color="5A5A5A", end_color="5A5A5A", fill_type="solid") # Gris foncé pour les préférences de -2
    for i in range(nbr_prof):
        for j in range(nbr_cours):
            if pref_prof[i][j] == -1 and y_output[i,j] == 1:
                sheet.cell(row=i+2, column=j+2).fill = couleur_m1 # +2 à cause des en-têtes
            elif pref_prof[i][j] == -2 and y_output[i,j] == 1:
                sheet.cell(row=i+2, column=j+2).fill = couleur_m2 # +2 à cause des en-têtes

    # Styles de bordure pour les cellules
    thin_border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin'))
        
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            cell.border = thin_border

    # Colonne (% cours) en % avec 1 décimale
    col_liberation = df_tache.columns.get_loc('% cours') + 1
    for row in range(2, nbr_prof + 2):
        cell = sheet.cell(row=row, column=col_liberation)
        cell.number_format = '0.0%'

    # Ajout du résumé sous le tableau
    resume_row = nbr_prof + 2 # Première ligne après les données des profs
    
    # Titres des cours
    for j in range(nbr_cours):
        cell = sheet.cell(row=resume_row, column=j+2, value=str(list_cours[j]))
        cell.border = thin_border

    # "Commande" - nombre de groupes à assigner par cours
    sheet.cell(row=resume_row+1, column=1, value="Commande").border = thin_border
    for j in range(nbr_cours):
        cell = sheet.cell(row=resume_row+1, column=j+2, value=int(groupes_cours[j]))
        cell.border = thin_border
    
    # "Assignés" - nombre de groupes réellement assignés par cours
    sheet.cell(row=resume_row+2, column=1, value="Assignés").border = thin_border
    for j in range(nbr_cours):
        assigned = int(sum(x_output[i,j] for i in range(nbr_prof)))
        cell = sheet.cell(row=resume_row+2, column=j+2, value=assigned)
        cell.border = thin_border
    
    # Ligne "Nbr. Étud." - nombre d'étudiants par cours
    sheet.cell(row=resume_row+3, column=1, value="Nbr. Étud.").border = thin_border
    for j in range(nbr_cours):
        cell = sheet.cell(row=resume_row+3, column=j+2, value=etudiants_cours[j])
        cell.border = thin_border
    
    # Ligne "Per/sem" - nombre de périodes par semaine par cours
    sheet.cell(row=resume_row+4, column=1, value="Pér./sem.").border = thin_border
    for j in range(nbr_cours):
        cell = sheet.cell(row=resume_row+4, column=j+2, value=int(periodes_cours[j]))
        cell.border = thin_border

    # Ajout des statistiques sur les préférences (2 colonnes à droite du tableau principal)
    stats_col = len(df_tache.columns) + 2
    
    # Statistiques
    stats = [
        ("Moy. pénal.", round(np.mean(pref_pond_heure_values), 2)),
        ("É.-t. pénal.", round(np.std(pref_pond_heure_values), 2)),
        ("Min. préf", round(np.min(pref_pond_heure_values), 2)),
        ("", ""), # Ligne vide
        ("Prof à 1 prép", sum(1 for val in nbr_prep_values if val == 1)),
        ("Prof à 4 gr.", sum(1 for val in nbr_gr_values if val >= 4)),
        ("Prof à 15 h.", sum(1 for val in nbr_heures_values if val >= 15)),
        ("Tot. satisfaits", sum(1 for val in pref_pond_heure_values if val == 1)),
        ("", ""), # Ligne vide
        ("CI Moy.", round(np.mean(ci_pleine_values), 2)),
        ("CI É.-t.", round(np.std(ci_pleine_values), 2)),
        ("CI Min", round(np.min(ci_pleine_values), 2)),
        ("CI Max", round(np.max(ci_pleine_values), 2)),
        ("sCI/40", round(sum(ci_values)/40, 2)),
        ("CI Année Moy.", round(np.mean(ci_annee_values), 2)),
        ("CI Année É.-t.", round(np.std(ci_annee_values), 2)),
        ("CI Année Min", round(np.min(ci_annee_values), 2)),
        ("CI Année Max", round(np.max(ci_annee_values), 2)),
    ]
    
    # Écrire les statistiques dans la feuille Excel
    for i, (label, value) in enumerate(stats, start=2):
        if label: # Ignorer les lignes vides
            sheet.cell(row=i, column=stats_col, value=label).border = thin_border
            sheet.cell(row=i, column=stats_col + 1, value=value).border = thin_border


    # Message d'avertissement si la tâche n'est pas optimale
    if prob.sol_status == pl.LpSolutionIntegerFeasible: # Tâche réalisable mais pas optimale
        warning_cell = sheet.cell(row=resume_row + 1, column=stats_col - 6, value="Attention : tâche non optimale")
        warning_cell.font = Font(color="FF0000", bold=True) # Rouge et gras

    # Centrer le contenu de la feuille de calcul
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Écrire les titres des cours à la verticale (dans la première ligne)
    for cell in sheet[1]:
        cell.alignment = Alignment(textRotation=90, horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 80
    
    # Écrire les titres des cours à la verticale (dans le résumé)
    for j in range(nbr_cours):
        cell = sheet.cell(row=resume_row, column=j+2)
        cell.alignment = Alignment(textRotation=90, horizontal="center", vertical="center")
    sheet.row_dimensions[resume_row].height = 60

    # Auto-ajustement de la largeur des colonnes (exclure la première ligne avec texte vertical)
    for column in sheet.columns:
        column_letter = column[0].column_letter
        # Calculer la largeur basée sur le contenu des cellules
        max_length = max(
            (len(str(cell.value)) for cell in column[1:] if cell.value),
            default=0
        )
        # Limiter la largeur: minimum 8, maximum 15 pour éviter les colonnes trop larges
        sheet.column_dimensions[column_letter].width = max(min(max_length + 2, 15), 8)
    
    # Affiche la préférence moyenne et min dans le terminal
    print()
    print(f"Tâche {numero_tache + 1}")
    print(f"Préférence moyenne: {np.mean(pref_pond_heure_values):.4f}")
    print(f"Préférence minimum: {np.min(pref_pond_heure_values):.2f}")
    print()

#################################################
# Résolution du problème de programmation linéaire
#################################################

#-----------------------------------------------
# Afficher un message d'information
#-----------------------------------------------

SEP = "-" * 95
def aff_msg(msg):
    print(f"\n{SEP}\n\n{msg}\n\n{SEP}")

#-----------------------------------------------
# Sélection et paramètres du solveur
#-----------------------------------------------

# Highs utilisé par défaut - Attention à installer les dépendances nécessaires, python -m pip install pulp[highs]
solver = pl.HiGHS(timeLimit=solver_time_limit, log_to_console = print_solver_log, log_file = solver_log_file, gapRel = solver_gap)

#-----------------------------------------------
# Résolution et création du fichier de sortie Excel
#-----------------------------------------------

# Vérifier si le fichier de sortie existe déjà et numéroter le nouveau si nécessaire
result_file = os.path.join(os.getcwd(), output_file + '.xlsx')
nbr_file = 1
while os.path.exists(result_file):
    result_file = os.path.join(os.getcwd(), f"{output_file}_{nbr_file}.xlsx")
    nbr_file += 1

aff_msg(f"Le résultat sera enregistré dans le fichier Excel : {os.path.basename(result_file)}")

with pd.ExcelWriter(result_file, engine="openpyxl") as writer:

    def abort(msg): # Affiche un message et arrête le processus
        aff_msg(msg)
        writer._save = lambda: None
        raise SystemExit(0)

    # Si l'optimisation lexicographique est activée, on maximise d'abord la préférence minimale
    if lexicographic_optimization:
        if min_lexico != 0: # Si on veut que la préférence minimale soit fixée manuellement à au moins min_lexico
            prob += min_pref_value >= min_lexico # Maintient le minimum défini manuellement

        else:
            prob += min_pref_value # Le minimum devient la priorité de l'optimisation (remplace l'ancienne fonction objectif)
            prob.solve(solver) # Résoudre afin de maximiser le minimum
            
            if prob.sol_status not in (pl.LpSolutionOptimal, pl.LpSolutionIntegerFeasible):
                abort("Aucune solution réalisable trouvée pour maximiser la préférence minimale. Arrêt du processus.")
            
            if prob.sol_status == pl.LpSolutionIntegerFeasible:
                aff_msg("Attention: le solveur n'a pas eu le temps de trouver le minimum optimal.")

            best_min = pl.value(prob.objective) # Sauvegarder le minimum optimal pour la suite

            if probleme_exact: # Inutile de l'afficher pour le problème approché, cette valeur n'a alors pas de signification satisfaisante
                if prob.sol_status == pl.LpSolutionIntegerFeasible:
                    aff_msg(f"Le meilleur minimum trouvé dans le temps imparti est : {best_min:.2f}")
                else:
                    aff_msg(f"Il n'existe aucune tâche pour laquelle la préférence minimale est supérieure à : {best_min:.2f}")

            prob += objective_fct # Réintroduire la fonction objectif originale qui maximise la moyenne des préférences
            prob += min_pref_value >= best_min - abs(best_min) * 0.005 # Ajouter comme contrainte de maintenir le minimum optimum, avec une petite tolérance

    # Résolution et sauvegarde de la première tâche
    prob.solve(solver)

    if prob.sol_status not in (pl.LpSolutionOptimal, pl.LpSolutionIntegerFeasible):
        abort("Aucune solution réalisable trouvée pour la tâche principale. Arrêt du processus.")

    if prob.sol_status == pl.LpSolutionIntegerFeasible:
        aff_msg("Attention : il ne s'agit pas d'une solution optimale.\nIl est déconseillé de générer des tâches alternatives à partir d'une tâche non optimale.")

    best_obj = pl.value(prob.objective) # Valeur optimale de la fonction objectif
    save_excel_file(writer, 0) # Sauvegarde la feuille "1"
    writer.book.save(result_file) # Sauvegarde du fichier Excel

    # Génération des tâches alternatives si demandé
    if nbr_taches_alternatives > 0: # Utilise une stratégie d'exclusion: exclure la solution précédente pour trouver une solution alternative différente.
        # Fonction pour exclure une solution donnée en y
        def exclure_y(prev_y_dict): # Force au moins un des professeurs à "perdre" au moins un des cours précédemment attribués. Gagner un cours ne suffit pas. Calculatoirement plus simple pour le solveur et force des changements plus importants.
            s1 = [(i, j) for (i, j), val in prev_y_dict.items() if val == 1]
            return pl.lpSum(y[(i, j)] for (i, j) in s1) <= len(s1) - diversite_factor
        
        # Génération des tâches alternatives
        for k in range(nbr_taches_alternatives):
            # Pas besoin de chercher une meilleure solution que la précédente, on sait qu'elle n'existe pas (en tout cas pour une solution optimale)
            obj_constr = objective_fct <= best_obj + 1e-6
            prob += obj_constr

            # Sauvegarder les valeurs de y de la solution précédente afin de l'exclure des tâches alternatives
            prev_y = {(i, j): get_val(y[(i, j)]) for i in range(nbr_prof) for j in range(nbr_cours)}

            # Exclure cette solution des tâches alternatives
            prob += exclure_y(prev_y)
            
            # Résolution de la tâche alternative
            prob.solve(solver)

            if prob.sol_status == pl.LpSolutionOptimal:
                best_obj = pl.value(prob.objective) # Valeur optimale de la fonction objectif pour la tâche courante. Inutile de la sauvegarder si elle n'est pas optimale, autant réutiliser la valeur précédente
            elif prob.sol_status == pl.LpSolutionIntegerFeasible:
                aff_msg("Attention : il ne s'agit pas d'une solution optimale.")

            # Suppression de la contrainte <= best_obj sur la fonction objectif (sera réintroduite pour la tâche alternative suivante)
            prob.constraints.pop(obj_constr.name, None)

            # Exportation de la solution alternative comme nouvelle feuille
            if prob.sol_status in (pl.LpSolutionOptimal, pl.LpSolutionIntegerFeasible):
                save_excel_file(writer, k + 1) # Feuille "2", "3", etc.
                writer.book.save(result_file) # Sauvegarde du fichier Excel
            else:
                aff_msg("Aucune solution réalisable trouvée pour cette tâche alternative. Arrêt de la génération de tâches alternatives.")
                break

    # On réordonne les feuilles dans le fichier Excel en fonction de la préférence moyenne, puis de la préférence minimum, puis de l'écart-type des préférences
    wb = writer.book
    if not wb.sheetnames: # Si aucune feuille n'a été créée, on affiche un message et on empêche la sauvegarde d'un classeur vide
        abort("Aucun résultat à sauvegarder.\nUn fichier Excel vide a été créé.")

    else:
        mesures = []
        for name in wb.sheetnames:
            ws = wb[name]
            # Les statistiques sont dans la dernière colonne
            stats_val_col = ws.max_column
            # On récupère les statistiques.
            moy_feuille = round(float(ws.cell(row=2, column=stats_val_col).value), 2)
            ecart_feuille = round(float(ws.cell(row=3, column=stats_val_col).value), 2)
            min_feuille = round(float(ws.cell(row=4, column=stats_val_col).value), 2)
            mesures.append((name, moy_feuille, min_feuille, ecart_feuille))

        # Tri des feuilles
        ordre = sorted(mesures, key=lambda t: (-t[1], -t[2], t[3]))
        ordre_nom_feuille = [t[0] for t in ordre]
        wb._sheets = [wb[s] for s in ordre_nom_feuille]

        # On renomme les feuilles pour que la première soit "1", la deuxième "2", etc.
        # Utiliser un préfixe temporaire pour éviter les conflits de noms lors du renommage
        for idx, sheet in enumerate(wb.worksheets):
            sheet.title = f"_tmp_{idx + 1}"
        for idx, sheet in enumerate(wb.worksheets):
            sheet.title = str(idx + 1)
        wb.active = 0