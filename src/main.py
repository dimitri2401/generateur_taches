#################################################
# Générateur de tâches - v0.3
#################################################

# Imports standards
import os
import sys

# Autres imports
import highspy
import numpy as np
import pandas as pd
import pulp as pl
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from ruamel.yaml import YAML


#################################################
# Log
#################################################

# Affichage en temps réel dans le terminal du gui
sys.stdout.reconfigure(line_buffering=True)


# Redirection de la sortie du terminal vers le fichier log
class Tee:
    def __init__(self, filepath):
        self.console = sys.stdout
        self.file = open(filepath, "w", encoding="utf-8")

    def write(self, message):
        self.console.write(message)
        self.file.write(message)

    def flush(self):
        self.console.flush()
        self.file.flush()


#################################################
# Fonction pour afficher un message d'information dans le terminal
#################################################

SEP = "-" * 95


def aff_msg(msg):
    print(f"\n{SEP}\n\n{msg}\n\n{SEP}\n")


#################################################
# Fonction principale
#################################################


def main():

    aff_msg("Générateur de tâches - v0.3")

    #################################################
    # Chargement des paramètres depuis le fichier YAML
    #################################################

    ryaml = YAML()
    with open("config.yaml", "r", encoding="utf-8") as f:
        config = ryaml.load(f)

    # -----------------------------------------------
    # Gestion des fichiers
    # -----------------------------------------------

    # Fichiers d'entrée/sortie
    input_file = config["fichiers"]["entree"]
    output_file = config["fichiers"]["sortie"]

    # -----------------------------------------------
    # Paramètres d'optimisation
    # -----------------------------------------------

    # Pré-optimisations
    ordre_preopti = config["optimisation"]["preoptimisations"]
    min_lexico = config["optimisation"]["seuil_prof_min"]
    nbr_prof_satisf = config["optimisation"]["seuil_nbr_satisfaits"]

    # Ne pas autoriser les préférences négatives
    disable_negative_preferences = config["optimisation"]["exclure_preferences_negatives"]

    # Nombre maximum d'heures de cours par prof
    max_hours = config["optimisation"]["nbr_heures_max"]

    # -----------------------------------------------
    # Paramètres pour la génération de tâches alternatives
    # -----------------------------------------------

    # Nombre de tâches supplémentaires à générer
    nbr_taches_alternatives = config["taches_alternatives"]["nbr_taches_alternatives"]
    # Facteur de diversité
    diversite_facteur = config["taches_alternatives"]["diversite"]

    # -----------------------------------------------
    # Log
    # -----------------------------------------------

    # Sauvegarder le log dans un fichier
    solver_log_file = config["log"]["fichier_log"] if config["log"]["sauvegarder_log"] else None
    if solver_log_file:
        logger = Tee(solver_log_file)
        sys.stdout = logger
        sys.stderr = logger

    # -----------------------------------------------
    # Paramètres du solveur - Non exposés dans le fichier de configuration YAML
    # -----------------------------------------------

    # Limite de temps pour le solveur en secondes
    solver_time_limit = 1000
    # Tolérance pour le solveur
    solver_gap = 0.0001

    #################################################
    # Chargement des données depuis le fichier Excel
    #################################################

    # Fonction de chargement des données depuis le fichier Excel
    def load_data(tache_file):
        if not os.path.exists(tache_file):
            print("Erreur : le fichier d'entrée n'existe pas.")
            raise SystemExit(1)

        # Charge les feuilles du fichier Excel de paramètres au format xls
        dfs = pd.read_excel(tache_file, sheet_name=["PROF", "COURS", "PREF", "MAX_NB_GR", "ATTRIB_PREALABLE"], header=None, engine="xlrd")  # Changer pour openpyxl si le fichier est en .xlsx

        df_prof = dfs["PROF"]
        # Emplacement des infos sur la page PROF (attention l'indexation commence à 0)
        CELL_NBR_PROF = (0, 0)  # Cellule contenant le nombre de professeurs
        COL_NOM_PROF = 0  # Colonne contenant les noms des professeurs
        COL_LIBERATION = 1  # Colonne contenant le % de tâche travaillée
        COL_CI_CIBLE = 2  # Colonne contenant la CI cible
        COL_TOL_INF = 3  # Colonne contenant la CI min
        COL_TOL_SUP = 4  # Colonne contenant la CI max
        COL_PREP_MIN = 5  # Colonne contenant le nombre minimum de préparations
        COL_PREP_MAX = 6  # Colonne contenant le nombre maximum de préparations
        COL_GR_MAX = 7  # Colonne contenant le nombre maximum de groupes
        COL_CI_ANT = 10  # Colonne contenant la CI antérieure

        # Extraction des données de la page PROF
        nbr_prof = int(df_prof.iat[CELL_NBR_PROF])
        list_prof = df_prof.iloc[1 : nbr_prof + 1, COL_NOM_PROF].astype(str).tolist()
        liberation_prof = df_prof.iloc[1 : nbr_prof + 1, COL_LIBERATION].astype(float).tolist()
        ci_min = (df_prof.iloc[1 : nbr_prof + 1, COL_CI_CIBLE] - df_prof.iloc[1 : nbr_prof + 1, COL_TOL_INF]).astype(float).tolist()
        ci_max = (df_prof.iloc[1 : nbr_prof + 1, COL_CI_CIBLE] + df_prof.iloc[1 : nbr_prof + 1, COL_TOL_SUP]).astype(float).tolist()
        prep_min = df_prof.iloc[1 : nbr_prof + 1, COL_PREP_MIN].astype(int).tolist()
        prep_max = df_prof.iloc[1 : nbr_prof + 1, COL_PREP_MAX].astype(int).tolist()
        gr_max = df_prof.iloc[1 : nbr_prof + 1, COL_GR_MAX].astype(int).tolist()
        ci_ant = df_prof.iloc[1 : nbr_prof + 1, COL_CI_ANT].astype(float).tolist()

        df_cours = dfs["COURS"]
        # Emplacement des infos sur la page COURS (attention l'indexation commence à 0)
        CELL_NBR_COURS = (0, 0)  # Cellule contenant le nombre de cours
        COL_NOM_COURS = 0  # Colonne contenant les noms des cours
        COL_PERIODES = 1  # Colonne contenant le nombre de périodes
        COL_GROUPES = 2  # Colonne contenant le nombre de groupes
        COL_ETUDIANTS = 3  # Colonne contenant le nombre d'étudiants

        # Extraction des données de la page COURS
        nbr_cours = int(df_cours.iat[CELL_NBR_COURS])
        list_cours = df_cours.iloc[1 : nbr_cours + 1, COL_NOM_COURS].astype(str).tolist()
        periodes_cours = df_cours.iloc[1 : nbr_cours + 1, COL_PERIODES].astype(int).tolist()
        groupes_cours = df_cours.iloc[1 : nbr_cours + 1, COL_GROUPES].astype(int).tolist()
        etudiants_cours = df_cours.iloc[1 : nbr_cours + 1, COL_ETUDIANTS].astype(int).tolist()

        # Extraction des préférences, nombres max de groupes et attributions préalables
        def lire_valeurs(sheet_name, default_value, sheet_label):
            df_sheet = dfs[sheet_name]
            # On ignore la première ligne et la première colonne
            data = df_sheet.iloc[1:, 1:].fillna(default_value)
            try:  # On s'assure que les données sont bien des nombres
                return data.apply(pd.to_numeric, errors="raise").astype(int)
            except Exception as e:
                print(f"Présence d'une valeur erronée dans la feuille {sheet_label} : {e}.")
                raise SystemExit(1)

        pref_data = lire_valeurs("PREF", 2, "PREF")
        max_nb_gr_data = lire_valeurs("MAX_NB_GR", 5, "MAX_NB_GR")
        attrib_preal_data = lire_valeurs("ATTRIB_PREALABLE", 0, "ATTRIB_PREALABLE")

        pref_prof = pref_data.iloc[:nbr_prof, :nbr_cours].to_numpy(dtype=int)  # On n'utilise pas .tolist() maintenant afin de pouvoir pré-calculer les préférences pénalisées au carré avec numpy plus tard
        max_nb_gr = max_nb_gr_data.iloc[:nbr_prof, :nbr_cours].to_numpy(dtype=int).tolist()
        attrib_preal = attrib_preal_data.iloc[:nbr_prof, :nbr_cours].to_numpy(dtype=int).tolist()

        return {
            "professors": {
                "list": list_prof,
                "nbr": nbr_prof,
                "liberation": liberation_prof,
                "ci_min": ci_min,
                "ci_max": ci_max,
                "prep_min": prep_min,
                "prep_max": prep_max,
                "gr_max": gr_max,
                "ci_ant": ci_ant,
            },
            "courses": {"list": list_cours, "nbr": nbr_cours, "periodes": periodes_cours, "groupes": groupes_cours, "etudiants": etudiants_cours},
            "preferences": {"prof": pref_prof, "max_nb_gr": max_nb_gr, "attrib_preal": attrib_preal},
        }

    # Chargement des données
    data = load_data(input_file)

    # Extraction des données
    nbr_prof = data["professors"]["nbr"]
    list_prof = data["professors"]["list"]
    liberation_prof = data["professors"]["liberation"]
    ci_min = data["professors"]["ci_min"]
    ci_max = data["professors"]["ci_max"]
    prep_min = data["professors"]["prep_min"]
    prep_max = data["professors"]["prep_max"]
    gr_max = data["professors"]["gr_max"]
    ci_ant = data["professors"]["ci_ant"]

    nbr_cours = data["courses"]["nbr"]
    list_cours = data["courses"]["list"]
    periodes_cours = data["courses"]["periodes"]
    groupes_cours = data["courses"]["groupes"]
    etudiants_cours = data["courses"]["etudiants"]

    pref_prof = data["preferences"]["prof"]
    max_nb_gr = data["preferences"]["max_nb_gr"]
    attrib_preal = data["preferences"]["attrib_preal"]

    # Validations
    for i in range(nbr_prof):
        for j in range(nbr_cours):
            # Validation des préférences
            if pref_prof[i, j] < -2 or pref_prof[i, j] > 2:
                print(f"Erreur : la préférence de {list_prof[i]} pour le cours {list_cours[j]} ({pref_prof[i, j]}) n'est pas dans l'intervalle [-2, 2].")
                raise SystemExit(1)

            # Validation des nombres de groupes max
            if max_nb_gr[i][j] < 0:
                print(f"Erreur : le nombre maximum de groupes de {list_prof[i]} pour le cours {list_cours[j]} ({max_nb_gr[i][j]}) est négatif.")
                raise SystemExit(1)

            # Validation des attributions préalables
            nb_groupes = attrib_preal[i][j]
            if nb_groupes < 0:
                print(f"Erreur : l'attribution préalable de {list_prof[i]} pour le cours {list_cours[j]} ({nb_groupes}) est négative.")
                raise SystemExit(1)

            if nb_groupes > max_nb_gr[i][j]:
                print(f"Erreur : l'attribution préalable de {list_prof[i]} pour le cours {list_cours[j]} ({nb_groupes}) dépasse son maximum souhaité ({max_nb_gr[i][j]}).")
                raise SystemExit(1)

            if disable_negative_preferences and pref_prof[i, j] < 0 and nb_groupes > 0:
                print(f"Erreur : l'attribution préalable de {list_prof[i]} pour le cours {list_cours[j]} ({nb_groupes}) entre en conflit avec sa préférence négative et l'option d'exclure les préférences négatives.")
                raise SystemExit(1)

    # Validation des attributions préalables et du nombre de groupes disponibles pour chaque cours
    for j in range(nbr_cours):
        total_attrib = sum(attrib_preal[i][j] for i in range(nbr_prof))
        if total_attrib > groupes_cours[j]:
            print(f"Erreur : les attributions préalables pour le cours {list_cours[j]} ({total_attrib}) dépassent le nombre de groupes disponibles ({groupes_cours[j]}).")
            raise SystemExit(1)

    #################################################
    # Calcul des tâches valides pour chaque professeur
    #################################################

    # -----------------------------------------------
    # Calcul des pénalités au carré
    # -----------------------------------------------

    # Pré-calcul des préférences pénalisées au carré pour tous les (i,j)
    pref_matrix = np.array(1 - (2 - pref_prof) ** 2 / 12 - (2 - pref_prof) / 6, dtype=float).tolist()

    # -----------------------------------------------
    # Calcul de la CI
    # -----------------------------------------------

    def calcul_ci(h_prep, h_cours, nes, pes, nbr_prep):
        if nbr_prep == 3:
            facteur_prep = 1.1
        elif nbr_prep >= 4:
            facteur_prep = 1.75
        else:
            facteur_prep = 0.9

        ci = facteur_prep * h_prep + 1.2 * h_cours
        if nes >= 75:
            ci += 0.01 * nes
            if nes >= 161:
                ci += 0.1 * (nes - 160) ** 2
        if pes >= 416:
            ci += 0.04 * 415 + 0.07 * (pes - 415)
        else:
            ci += 0.04 * pes
        return ci

    # -----------------------------------------------
    # Calcul des tâches valides pour chaque professeur
    # -----------------------------------------------

    # Tolérance pour les comparaisons de nombres flottants afin de ne pas manquer de tâches valides
    TOLERANCE = 1e-9

    # Nombre de groupes encore disponibles après déduction des attributions préalables
    groupes_dispo = [groupes_cours[j] - sum(attrib_preal[k][j] for k in range(nbr_prof)) for j in range(nbr_cours)]

    # Calcul pour le prof i des tâches valides respectant les contraintes de nombre de groupes, de nombre préparations et de CI
    def calcul_taches_prof(i):
        # Indices des cours que le prof i n'a pas exclus
        if disable_negative_preferences:
            cours_indices = [j for j in range(nbr_cours) if max_nb_gr[i][j] > 0 and pref_prof[i, j] >= 0]
        else:
            cours_indices = [j for j in range(nbr_cours) if max_nb_gr[i][j] > 0]

        # Tri des cours par nombre de périodes décroissant pour explorer d'abord les cours les plus longs et potentiellement finir plus vite la recherche
        cours_indices.sort(key=lambda j: periodes_cours[j], reverse=True)

        # Attributions préalables
        preal_distribution = [attrib_preal[i][j] for j in range(nbr_cours)]
        preal_k = sum(preal_distribution)

        # On commence à k = preal_k car on ne peut pas attribuer moins de groupes que ceux imposés par attrib_preal
        start_k = max(1, preal_k)

        # Si les attributions préalables dépassent le nombre de groupes possibles, on retourne une liste vide
        # Non nécessaire car on a déjà validé les attributions préalables précédemment
        # for j in range(nbr_cours):
        #     if preal_distribution[j] > min(max_nb_gr[i][j], groupes_cours[j]):
        #         return []

        groupes_actuels = list(preal_distribution)  # On copie la distribution initiale afin de pouvoir la modifier dans la suite

        # Calcul des informations initiales liées aux attributions préalables afin de démarrer la récursion
        initial_nbr_prep = sum(1 for j in range(nbr_cours) if groupes_actuels[j] > 0)
        initial_heures_prep = sum(periodes_cours[j] for j in range(nbr_cours) if groupes_actuels[j] > 0)
        initial_heures_cours = sum(periodes_cours[j] * groupes_actuels[j] for j in range(nbr_cours))
        initial_nes = sum(etudiants_cours[j] * groupes_actuels[j] for j in range(nbr_cours))
        initial_pes = sum(periodes_cours[j] * etudiants_cours[j] * groupes_actuels[j] for j in range(nbr_cours))

        # Calcul de la CI pour les attributions préalables
        initial_ci = calcul_ci(initial_heures_prep, initial_heures_cours, initial_nes, initial_pes, initial_nbr_prep)

        # Si les attributions préalables dépassent les contraintes, on retourne une liste vide
        if initial_ci > ci_max[i] + TOLERANCE or initial_heures_cours > max_hours or initial_nbr_prep > prep_max[i] or preal_k > gr_max[i]:
            return []

        # Calcul de la préférence initiale pour les attributions préalables
        initial_pref_score = sum(pref_matrix[i][j] * groupes_actuels[j] * periodes_cours[j] for j in cours_indices)

        # Calcul du nombre de cours éligibles pour le prof i
        nb_cours_eligibles = len(cours_indices)

        # Calcul du nombre maximum de groupes de chaque cours que l'on peut encore ajouter pour chaque cours
        ajouts_max_par_cours = [max(0, min(max_nb_gr[i][j] - preal_distribution[j], groupes_dispo[j])) for j in cours_indices]

        # Liste pour stocker les tâches valides
        taches_valides = []

        # Fonction pour enregistrer une tâche valide
        def enregistrer_tache(preference, heures_cours, ci, nb_prep):
            taches_valides.append(
                {
                    "gr_cours": groupes_actuels.copy(),  # Nombre de groupes pour chaque cours
                    "pref": round(preference / heures_cours, 4),  # On arrondit la préférence moyenne pondérée à 4 décimales pour éviter les problèmes de précision
                    "ci": round(ci, 2),
                    "heures_cours": heures_cours,
                    "nbr_prep": nb_prep,
                }
            )

        # Fonction d'exploration
        def explorer_taches(index, nb_groupes, nb_prep, heures_prep, heures_cours, nb_etudiants, periodes_etudiants, ci_actuelle, preference):
            # index : l'indice du cours actuel dans la liste des cours éligibles
            # nb_groupes : le nombre total de groupes attribués jusqu'à présent
            # nb_prep : le nombre total de préparations attribuées jusqu'à présent
            # heures_prep : le nombre total d'heures de préparation attribuées jusqu'à présent
            # heures_cours : le nombre total d'heures de cours attribuées jusqu'à présent
            # nb_etudiants : le nombre total d'étudiants attribués jusqu'à présent
            # periodes_etudiants : le nombre total de périodes-étudiants attribuées jusqu'à présent
            # ci_actuelle : la CI jusqu'à présent
            # preference : la préférence (pas encore divisée par le nombre d'heures) jusqu'à présent

            # Condition d'arrêt : on arrive à la fin de la liste des cours éligibles
            if index == nb_cours_eligibles:
                # Si la tâche actuelle respecte les contraintes, on l'ajoute à la liste des tâches valides. Pas besoin de bornes sup, elles sont déjà imposées par la récursion
                if nb_groupes >= start_k and nb_prep >= prep_min[i] and ci_actuelle >= ci_min[i] - TOLERANCE:
                    enregistrer_tache(preference, heures_cours, ci_actuelle, nb_prep)
                return

            # Nombre de cours qu'il reste à explorer dans cette branche
            cours_restants = nb_cours_eligibles - index

            # Si, en ajoutant une nouvelle préparation pour chaque cours restant on n'atteint pas le min de prep requis, on arrête
            if nb_prep + cours_restants < prep_min[i]:
                return

            # Si le prof a déjà atteint son nombre maximum de groupes ou d'heures, on arrête
            if nb_groupes >= start_k and (nb_groupes == gr_max[i] or heures_cours == max_hours):
                if nb_prep >= prep_min[i] and ci_actuelle >= ci_min[i] - TOLERANCE:
                    enregistrer_tache(preference, heures_cours, ci_actuelle, nb_prep)
                return

            # On récupère le prochain cours à explorer
            j = cours_indices[index]

            # On calcule le nombre maximum de groupes que l'on peut encore ajouter pour ce cours
            heures_restantes = max_hours - heures_cours
            max_groupes_heures = heures_restantes // periodes_cours[j]
            max_groupes_ajoutables = min(ajouts_max_par_cours[index], gr_max[i] - nb_groupes, max_groupes_heures)

            # Si le prof n'a pas d'attribution préalable pour ce cours et qu'il a déjà atteint son nombre maximum de préparations, on ne peut pas ajouter ce cours
            if preal_distribution[j] == 0 and nb_prep >= prep_max[i]:
                max_groupes_ajoutables = 0

            # On explore toutes les possibilités d'ajouts de groupes pour ce cours

            # 0 groupe ajouté pour ce cours, on passe au cours suivant
            explorer_taches(index + 1, nb_groupes, nb_prep, heures_prep, heures_cours, nb_etudiants, periodes_etudiants, ci_actuelle, preference)

            # 1 groupe ou plus
            for nb_groupes_ajoutes in range(1, max_groupes_ajoutables + 1):
                # Détermine si l'ajout de ces groupes pour ce cours donne une nouvelle préparation
                bool_nouvelle_prep = 1 if (preal_distribution[j] == 0) else 0

                # On calcule les nouvelles valeurs après l'ajout de ces groupes
                nouveau_nb_prep = nb_prep + bool_nouvelle_prep
                nouveau_heures_prep = heures_prep + (periodes_cours[j] * bool_nouvelle_prep)
                nouveau_heures_cours = heures_cours + (periodes_cours[j] * nb_groupes_ajoutes)
                nouveau_nb_etudiants = nb_etudiants + (etudiants_cours[j] * nb_groupes_ajoutes)
                nouveau_periodes_etudiants = periodes_etudiants + (periodes_cours[j] * etudiants_cours[j] * nb_groupes_ajoutes)
                nouveau_ci = calcul_ci(nouveau_heures_prep, nouveau_heures_cours, nouveau_nb_etudiants, nouveau_periodes_etudiants, nouveau_nb_prep)

                # Si la nouvelle CI dépasse la CI max, on arrête l'exploration pour ce cours
                if nouveau_ci > ci_max[i] + TOLERANCE:
                    break

                # Sinon on ajoute les groupes pour ce cours et on continue l'exploration
                groupes_actuels[j] += nb_groupes_ajoutes
                explorer_taches(index + 1, nb_groupes + nb_groupes_ajoutes, nouveau_nb_prep, nouveau_heures_prep, nouveau_heures_cours, nouveau_nb_etudiants, nouveau_periodes_etudiants, nouveau_ci, preference + (pref_matrix[i][j] * nb_groupes_ajoutes * periodes_cours[j]))

                # On retire les groupes ajoutés pour ce cours pour revenir à l'état précédent
                groupes_actuels[j] -= nb_groupes_ajoutes

        # Lancement de la récursion
        explorer_taches(0, preal_k, initial_nbr_prep, initial_heures_prep, initial_heures_cours, initial_nes, initial_pes, initial_ci, initial_pref_score)

        return taches_valides

    # On précalcule, pour chaque prof, les tâches valides
    aff_msg("Pré-calcul des tâches valides pour chaque professeur.")
    taches_departement = []  # Liste des tâches valides de l'ensemble des professeurs
    for i in range(nbr_prof):
        tache_i = calcul_taches_prof(i)

        # Validation: si aucune tâche valide n'existe pour un prof, on retourne une erreur et on arrête le programme
        if not tache_i:
            print(f"Erreur : aucune tâche valide n'existe pour {list_prof[i]}. Vérifiez les paramètres de la tâche.")
            raise SystemExit(1)

        if len(tache_i) >= 2:
            print(f"{list_prof[i]} : {len(tache_i)} tâches valides trouvées.")
        else:
            print(f"{list_prof[i]} : 1 tâche valide trouvée.")

        taches_departement.append(tache_i)

    #################################################
    # Définition du problème de programmation linéaire
    #################################################

    # -----------------------------------------------
    # Création du problème et des variables de décision
    # -----------------------------------------------

    # Création du problème
    prob = pl.LpProblem("probleme_tache", pl.LpMaximize)

    # Variables de décision est_attrib[(i, c)] : la tâche c est-elle attribuée au prof i (1 si oui, 0 sinon)
    est_attrib = pl.LpVariable.dicts("est_attrib", ((i, c) for i in range(nbr_prof) for c in range(len(taches_departement[i]))), cat="Binary")

    # -----------------------------------------------
    # Contraintes d'optimisation
    # -----------------------------------------------

    # Chaque prof doit recevoir exactement une tâche
    for i in range(nbr_prof):
        prob += pl.lpSum([est_attrib[(i, c)] for c in range(len(taches_departement[i]))]) == 1

    # Tous les groupes de tous les cours doivent être attribués
    for j in range(nbr_cours):
        prob += pl.lpSum([est_attrib[(i, c)] * taches_departement[i][c]["gr_cours"][j] for i in range(nbr_prof) for c in range(len(taches_departement[i]))]) == groupes_cours[j]

    # -----------------------------------------------
    # Fonction objectif principale
    # -----------------------------------------------

    # On maximise la somme des préférences moyennes pondérées
    objective_fct = pl.lpSum([est_attrib[(i, c)] * taches_departement[i][c]["pref"] for i in range(nbr_prof) for c in range(len(taches_departement[i]))])

    #################################################
    # Préparation des données de sortie
    #################################################

    # Variables de sortie
    x_output = np.zeros((nbr_prof, nbr_cours), dtype=int)  # Nombres de groupes pour chaque prof et chaque cours
    ci_output = np.zeros(nbr_prof, dtype=float)  # CI pour chaque prof
    hc_output = np.zeros(nbr_prof, dtype=int)  # Nombre d'heures de cours pour chaque prof
    nbr_prep_output = np.zeros(nbr_prof, dtype=int)  # Nombre de préparations pour chaque prof
    pref_output = np.zeros(nbr_prof, dtype=float)  # Préférence moyenne pondérée pour chaque prof

    def output_values():  # On remplit les variables de sortie à partir des valeurs des variables de décision
        x_output.fill(0)

        for i in range(nbr_prof):
            for c in range(len(taches_departement[i])):
                if round(est_attrib[(i, c)].varValue) == 1:  # round pour éviter les problèmes de précision du solveur
                    current_tache = taches_departement[i][c]

                    ci_output[i] = current_tache["ci"]
                    hc_output[i] = current_tache["heures_cours"]
                    nbr_prep_output[i] = current_tache["nbr_prep"]
                    pref_output[i] = current_tache["pref"]

                    for j in range(nbr_cours):
                        groups = current_tache["gr_cours"][j]
                        x_output[i, j] = groups
                    break  # Une seule tâche est attribuée par prof, on peut donc sortir de la boucle dès qu'on en trouve une

    #################################################
    # Exportation des résultats dans un fichier Excel
    #################################################

    # Fonction d'exportation des résultats dans un fichier Excel
    def save_excel_file(writer, numero_tache):  # Exportation des résultats dans un fichier Excel comme nouvelle feuille
        output_values()  # Valeurs de sortie à partir des variables de décision

        # Création du DataFrame
        df_tache = pd.DataFrame(x_output, columns=list_cours)

        # Infos
        df_tache.insert(0, "Prof.", list_prof)
        df_tache["% cours"] = liberation_prof
        df_tache["CI Cours"] = ci_output
        df_tache["CI Pleine"] = ci_output + 40 * (1 - np.array(liberation_prof))
        df_tache["CI Année"] = np.array(ci_ant) + df_tache["CI Pleine"]
        df_tache["Nbr. Prép."] = nbr_prep_output
        df_tache["Nbr. Groupes"] = x_output.sum(axis=1)
        df_tache["Nbr. Périodes"] = hc_output
        df_tache["Préf."] = np.round(pref_output, 2)

        # Nom de la feuille Excel. "1" pour la première, "2", "3", etc. pour les alternatives
        sheet_name = str(numero_tache + 1)

        # Mise en forme et enregistrement du fichier Excel
        df_tache.to_excel(writer, sheet_name=sheet_name, index=False)
        sheet = writer.sheets[sheet_name]

        # On définit le style de bordures
        if "thin_border" not in writer.book.style_names:
            thin_border = NamedStyle(name="thin_border")
            thin_border.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
            writer.book.add_named_style(thin_border)

        # On définit les couleurs et les alignements utilisés
        couleur_m1 = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        couleur_m2 = PatternFill(start_color="5A5A5A", end_color="5A5A5A", fill_type="solid")
        center_alignment = Alignment(horizontal="center", vertical="center")
        vertical_alignment = Alignment(textRotation=90, horizontal="center", vertical="center")

        # Nombre de colonnes dans le tableau principal
        num_cols = len(df_tache.columns)  # Nombre de colonnes dans le tableau principal

        # On récupère l'index de la colonne "% cours" pour appliquer le format de pourcentage
        col_liberation = df_tache.columns.get_loc("% cours") + 1

        # On collecte les informations du résumé
        resume_row = nbr_prof + 2
        resume_data = [("Titres", list_cours), ("Commande", groupes_cours), ("Assignés", [int(sum(x_output[i, j] for i in range(nbr_prof))) for j in range(nbr_cours)]), ("Nbr. Étud.", etudiants_cours), ("Pér./sem.", periodes_cours)]

        # On remplit le résumé
        for row_offset, (label, values) in enumerate(resume_data):
            sheet.cell(row=resume_row + row_offset, column=1, value=label).style = "thin_border"
            sheet.cell(row=resume_row + row_offset, column=1, value=label).alignment = center_alignment
            for j in range(nbr_cours):
                sheet.cell(row=resume_row + row_offset, column=j + 2, value=values[j]).style = "thin_border"
                sheet.cell(row=resume_row + row_offset, column=j + 2, value=values[j]).alignment = center_alignment

        # On collecte les statistiques sur les préférences. Elles sont calculées par rapport aux préférences réellement inscrites dans le fichier Excel (donc arrondies à 2 décimales) alors que les préférences sont à 4 décimales pour l'optimisation
        stats_col = len(df_tache.columns) + 2
        stats = [
            ("Moy. pénal.", round(df_tache["Préf."].mean(), 2)),
            ("É.-t. pénal.", round(df_tache["Préf."].std(ddof=0), 2)),
            ("Min. préf.", round(df_tache["Préf."].min(), 2)),
            ("", ""),
            ("Prof à 1 prép.", int((df_tache["Nbr. Prép."] == 1).sum())),
            ("Prof à 4+ gr.", int((df_tache["Nbr. Groupes"] >= 4).sum())),
            ("Prof à 15+ h.", int((df_tache["Nbr. Périodes"] >= 15).sum())),
            ("Tot. satisfaits", int((df_tache["Préf."] >= 0.9951).sum())),
            ("", ""),
            ("CI Moy.", round(df_tache["CI Pleine"].mean(), 2)),
            ("CI É.-t.", round(df_tache["CI Pleine"].std(ddof=0), 2)),
            ("CI Min.", round(df_tache["CI Pleine"].min(), 2)),
            ("CI Max.", round(df_tache["CI Pleine"].max(), 2)),
            ("sCI/40", round(ci_output.sum() / 40, 2)),
            ("CI Année Moy.", round(df_tache["CI Année"].mean(), 2)),
            ("CI Année É.-t.", round(df_tache["CI Année"].std(ddof=0), 2)),
            ("CI Année Min.", round(df_tache["CI Année"].min(), 2)),
            ("CI Année Max.", round(df_tache["CI Année"].max(), 2)),
        ]

        # On remplit les statistiques
        for i, (label, value) in enumerate(stats, start=2):
            if label:
                sheet.cell(row=i, column=stats_col, value=label).style = "thin_border"
                sheet.cell(row=i, column=stats_col, value=label).alignment = center_alignment
                sheet.cell(row=i, column=stats_col + 1, value=value).style = "thin_border"
                sheet.cell(row=i, column=stats_col + 1, value=value).alignment = center_alignment

        # Message d'avertissement si la tâche n'est pas optimale
        if prob.sol_status == pl.LpSolutionIntegerFeasible:
            warning_cell = sheet.cell(row=resume_row + 1, column=stats_col - 6, value="Attention : tâche non optimale")
            warning_cell.font = Font(color="FF0000", bold=True)  # En rouge et gras

        # On applique les bordures et l'alignement pour le tableau principal
        for row in sheet.iter_rows(min_row=1, max_row=nbr_prof + 1, min_col=1, max_col=num_cols):
            for cell in row:
                cell.style = "thin_border"
                cell.alignment = vertical_alignment if cell.row == 1 else center_alignment

        # On applique le format de pourcentage à la colonne "% cours"
        for r in range(2, nbr_prof + 2):
            sheet.cell(row=r, column=col_liberation).number_format = "0.0%"

        # On applique l'alignement vertical pour la ligne de résumé
        for c in range(2, nbr_cours + 2):
            sheet.cell(row=resume_row, column=c).alignment = vertical_alignment

        # On applique les couleurs pour les préférences négatives dans le tableau principal
        for i in range(nbr_prof):
            for j in range(nbr_cours):
                if x_output[i, j] > 0:
                    pref = pref_prof[i, j]
                    if pref == -1:
                        sheet.cell(row=i + 2, column=j + 2).fill = couleur_m1
                    elif pref == -2:
                        sheet.cell(row=i + 2, column=j + 2).fill = couleur_m2

        # On ajuste la hauteur de la première ligne et de la ligne de résumé
        sheet.row_dimensions[1].height = 80
        sheet.row_dimensions[resume_row].height = 60

        # Auto-ajustement de la largeur des colonnes (en excluant la première ligne avec texte vertical)
        for column in sheet.columns:
            column_letter = column[0].column_letter
            max_length = max((len(str(cell.value)) for cell in column[1:] if cell.value), default=0)
            sheet.column_dimensions[column_letter].width = max(min(max_length + 2, 15), 8)

        # On affiche la préférence moyenne et min dans le terminal, calculées à partir des préférences arrondies à 4 décimales issues de l'optimisation
        print()
        print(f"Tâche {numero_tache + 1}")
        print(f"Préférence moyenne : {pref_output.mean():.4f}")
        print(f"Préférence minimum : {pref_output.min():.2f}")
        print()

    #################################################
    # Résolution du problème de programmation linéaire
    #################################################

    # -----------------------------------------------
    # Paramètres du solveur HiGHS
    # -----------------------------------------------

    # On extrait les messages de HiGHS et les affiche dans le terminal
    def highs_log_callback(callback_type, message, data_out, data_in, user_data):
        print(message, end="")

    solver = pl.HiGHS(
        timeLimit=solver_time_limit,
        log_to_console=False,
        gapRel=solver_gap,
        callbacksToActivate=[highspy.cb.HighsCallbackType.kCallbackLogging],
        callbackTuple=(highs_log_callback, None),
    )

    # -----------------------------------------------
    # Résolution et création du fichier de sortie Excel
    # -----------------------------------------------

    # On vérifie si le fichier de sortie existe déjà et numéroter le nouveau si nécessaire
    result_file = os.path.join(os.getcwd(), output_file + ".xlsx")
    nbr_file = 1
    while os.path.exists(result_file):
        result_file = os.path.join(os.getcwd(), f"{output_file}_{nbr_file}.xlsx")
        nbr_file += 1

    aff_msg(f"Le résultat sera enregistré dans le fichier Excel : {os.path.basename(result_file)}")

    with pd.ExcelWriter(result_file, engine="openpyxl") as writer:

        def abort(msg):  # Affiche un message, sauvegarde le fichier et arrête le processus
            aff_msg(msg)
            writer._save = lambda: None
            raise SystemExit(0)

        # Objectif : maximiser le nombre de profs totalement satisfaits
        def phase_max_satisf():
            nonlocal prob

            total_satisf = pl.lpSum([est_attrib[(i, c)] for i in range(nbr_prof) for c in range(len(taches_departement[i])) if taches_departement[i][c]["pref"] >= 0.9951])

            if nbr_prof_satisf != 0:  # Si seuil fixé manuellement
                prob += total_satisf >= nbr_prof_satisf  # Maintient le seuil défini manuellement
            else:
                prob += total_satisf  # Maximiser le nombre de profs totalement satisfaits devient (temporairement) la priorité de l'optimisation
                prob.solve(solver)  # On résout pour cet objectif temporaire

                if prob.sol_status not in (pl.LpSolutionOptimal, pl.LpSolutionIntegerFeasible):
                    abort("Aucune solution réalisable trouvée pour maximiser les profs totalement satisfaits. Arrêt du processus.")

                if prob.sol_status == pl.LpSolutionIntegerFeasible:
                    aff_msg("Attention : le solveur n'a pas eu le temps de trouver le nombre optimal de profs totalement satisfaits.")

                best_count = round(pl.value(prob.objective))  # On sauvegarde le nombre optimal de profs totalement satisfaits pour la suite
                aff_msg(f"Nombre maximal de profs totalement satisfaits possible : {best_count}")

                prob += total_satisf >= best_count  # On force à maintenir le nombre de profs satisfaits optimal

        # Objectif : maximiser la préférence minimale parmi les profs
        def phase_lexicographic():
            nonlocal prob

            # Variable pour stocker la préférence minimale parmi les profs
            min_pref_value = pl.LpVariable("min_pref_value", lowBound=-1, upBound=1, cat="Continuous")

            for i in range(nbr_prof):
                prob += min_pref_value <= pl.lpSum([est_attrib[(i, c)] * taches_departement[i][c]["pref"] for c in range(len(taches_departement[i]))])

            if min_lexico != 0:  # Si seuil fixé manuellement
                prob += min_pref_value >= min_lexico  # Maintient le seuil défini manuellement

            else:
                prob += min_pref_value  # Le minimum devient (temporairement) la priorité de l'optimisation
                prob.solve(solver)  # On résout pour cet objectif temporaire

                if prob.sol_status not in (pl.LpSolutionOptimal, pl.LpSolutionIntegerFeasible):
                    abort("Aucune solution réalisable trouvée pour maximiser la préférence minimale. Arrêt du processus.")

                if prob.sol_status == pl.LpSolutionIntegerFeasible:
                    aff_msg("Attention : le solveur n'a pas eu le temps de trouver le minimum optimal.")

                best_min = pl.value(prob.objective)  # On sauvegarde le minimum optimal pour la suite

                if prob.sol_status == pl.LpSolutionIntegerFeasible:
                    aff_msg(f"Le meilleur minimum trouvé dans le temps imparti est : {best_min:.2f}")
                else:
                    aff_msg(f"Il n'existe aucune tâche pour laquelle la préférence minimale est supérieure à : {best_min:.2f}")

                prob += min_pref_value >= best_min - max(abs(best_min) * 0.01, 1e-3)  # On force à maintenir le minimum optimum, avec une petite tolérance (1%)

        # Exécution des phases de pré-optimisation dans l'ordre configuré dans le fichier YAML
        phases = {"prof_min": phase_lexicographic, "nbr_satisfaits": phase_max_satisf}
        for phase_name in ordre_preopti:
            phases[phase_name]()

        # On résout et sauvegarde la première tâche (hors pré-optimisations)
        prob += objective_fct  # On ajoute la fonction objectif principale (moyenne des préférences) au problème (elle remplace les fonctions objectifs temporaires des phases de pré-optimisation si elles ont été exécutées)
        prob.solve(solver)  # On résout le problème

        if prob.sol_status not in (pl.LpSolutionOptimal, pl.LpSolutionIntegerFeasible):
            abort("Aucune solution réalisable trouvée pour la tâche principale. Arrêt du processus.")

        if prob.sol_status == pl.LpSolutionIntegerFeasible:
            aff_msg("Attention : il ne s'agit pas d'une solution optimale.\nIl est déconseillé de générer des tâches alternatives à partir d'une tâche non optimale.")

        save_excel_file(writer, 0)  # On sauvegarde la feuille "1"
        writer.book.save(result_file)  # On sauvegarde le fichier Excel

        # Génération des tâches alternatives
        if nbr_taches_alternatives > 0:
            # Fonction pour exclure les solutions précédentes. On impose qu'au moins diversite_facteur profs abandonnent au moins un cours qu'ils enseignaient par rapport à toutes les solutions précédentes
            def exclure_sol():
                # On identifie, pour chaque prof, les cours qu'il enseignait (donc avec gr_cours > 0) dans la dernière solution générée
                cours_prec = {}
                for i in range(nbr_prof):
                    for c in range(len(taches_departement[i])):
                        if round(est_attrib[(i, c)].varValue) == 1:
                            cours_prec[i] = taches_departement[i][c]["gr_cours"]
                            break

                # On identifie les combinaisons qui impliquent l'abandon complet d'au moins un de ces cours
                combinaisons_abandon = [(i, c) for i in range(nbr_prof) for c in range(len(taches_departement[i])) if any(prec > 0 and actuel == 0 for prec, actuel in zip(cours_prec[i], taches_departement[i][c]["gr_cours"]))]

                # On force qu'au moins diversite_facteur profs abandonnent au moins un cours qu'ils enseignaient dans la dernière solution générée
                return pl.lpSum([est_attrib[(i, c)] for (i, c) in combinaisons_abandon]) >= diversite_facteur

            # Génération des tâches alternatives
            for k in range(nbr_taches_alternatives):
                # On exclut la solution précédente (cumulativement, donc toutes les solutions précédentes sont exclues)
                prob += exclure_sol()

                # On résout la tâche alternative
                prob.solve(solver)

                if prob.sol_status == pl.LpSolutionIntegerFeasible:
                    aff_msg("Attention : il ne s'agit pas d'une solution optimale.")

                # On exporte la solution alternative dans une nouvelle feuille
                if prob.sol_status in (pl.LpSolutionOptimal, pl.LpSolutionIntegerFeasible):
                    save_excel_file(writer, k + 1)  # Feuille "2", "3", etc.
                    writer.book.save(result_file)  # Sauvegarde du fichier Excel
                else:
                    aff_msg("Aucune solution réalisable trouvée pour cette tâche alternative. Arrêt de la génération de tâches alternatives.")
                    break

        # On réordonne les feuilles dans le fichier Excel en fonction de la préférence moyenne, puis de la préférence minimum, puis de l'écart-type des préférences
        wb = writer.book
        if not wb.sheetnames:  # Si aucune feuille n'a été créée, on affiche un message, on empêche la sauvegarde d'un classeur vide et on quitte
            abort("Aucun résultat à sauvegarder.\nUn fichier Excel vide a été créé.")

        # Sinon on récupère les statistiques de chaque feuille pour les trier
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

        # On trie les feuilles selon la préférence moyenne
        ordre = sorted(mesures, key=lambda t: (-t[1], -t[2], t[3]))
        ordre_nom_feuille = [t[0] for t in ordre]
        wb._sheets = [wb[s] for s in ordre_nom_feuille]

        # On renomme les feuilles pour que la première soit "1", la deuxième "2", etc.
        # On utilise un préfixe temporaire pour éviter les conflits de noms lors du renommage
        for idx, sheet in enumerate(wb.worksheets):
            sheet.title = f"_tmp_{idx + 1}"
        for idx, sheet in enumerate(wb.worksheets):
            sheet.title = str(idx + 1)
        wb.active = 0


if __name__ == "__main__":
    main()
