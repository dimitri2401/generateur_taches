## Paramètres de configuration du générateur de tâches

**Remarque :** les paramètres disponibles sont susceptibles d'être ajustés ou supprimés au fil des expérimentations.

### Fichier de configuration :

Ce document décrit les paramètres disponibles dans le fichier de configuration `config.yaml` du générateur de tâches et explique leur effet. Le fichier de configuration peut être ouvert avec n'importe quel éditeur de texte.

**Important :** ne supprimer aucune entrée du fichier de configuration.

### Fichiers d'entrée/sortie (`files`)
- **`input`** : nom du fichier contenant les paramètres de la tâche, avec son extension (p. ex. `tache.xls`). C'est le fichier Excel lu par le générateur, contenant les informations sur les professeurs, les cours, les préférences, les nombres de groupes, etc.
- **`output`** : nom du fichier de sortie, sans extension (p. ex. `resultat_tache`). Si le fichier existe déjà dans le répertoire (p. ex. en raison d'une exécution précédente du logiciel), le nouveau fichier sera renommé en lui ajoutant un numéro.

### Paramètres d'optimisation (`optimization`)

#### Fonction objectif

Il est possible d'influencer la fonction objectif à maximiser.

**Remarque** : la quantité \( pref(i,j) \) représente la préférence pénalisée au carré (aussi appelée ppc) :
$$
pref(i,j) = 1 - \frac{(2 - p)^2}{12} - \frac{(2 - p)}{6}
$$

- **`probleme_exact`** : `True`/`False`. Si `True`, la fonction suivante est utilisée comme objectif. C'est historiquement celle que le logiciel d'Éric utilisait comme métrique de qualité des tâches. Cette option donne la solution optimale au véritable problème de la tâche, mais rend la résolution potentiellement longue.
$$
\sum_{\text{prof } i} \frac{1}{\text{nbr d'heures enseignées par le prof \( i \)}} \sum_{\text{cours } j} \text{nbr de grp du cours \( j \) attribués} \times \text{nbr d'heures du cours \( j \)} \times pref(i , j)
$$

Si l'option précédente est désactivée, c'est une version approchée du problème qui est résolue. Il s'agit alors de maximiser la fonction suivante :
$$
\sum_{\text{prof } i} \sum_{\text{cours } j} \text{nbr de grp du cours \( j \) attribués} \times \text{nbr d'heures du cours \( j \)} \times pref(i , j)
$$

Ce problème simplifié aboutit à des solutions essentiellement équivalentes au problème exact, même si, parfois, elles ne sont pas totalement optimales. En revanche, sa résolution est généralement beaucoup plus rapide (environ un ordre de grandeur sur les tâches testées).

Si l'option `probleme_exact` est désactivée, il est possible de simuler (partiellement) l'effet des libérations dans la version simplifiée de la fonction objectif :

- **`liberation_enabled`** : `True`/`False`. Si `True`, tente de simuler l'effet des libérations dans le problème approché. Les professeurs libérés ayant moins de groupes ou d'heures, leur préférence a parfois un peu moins de poids dans le total. **Important :** ne peut pas être activé dans le cas du problème exact.

#### Maximisation du minimum des préférences individuelles

Maximiser la fonction objectif ne garantit pas nécessairement que l'ensemble des professeurs aient des tâches satisfaisantes. Il ne s'agit en effet que de maximiser la moyenne des préférences.

Afin de s'assurer que la préférence minimale parmi l'ensemble des professeurs ne soit pas trop basse, il est possible d'utiliser l'une des deux méthodes suivantes.

Optimisation lexicographique :

- **`lexicographic_optimization`** : `True`/`False`. Si `True`, active une optimisation en deux étapes. Premièrement, on cherche, parmi toutes les tâches possibles, celle pour laquelle la préférence minimale parmi l'ensemble des professeurs est maximale (sauf si une valeur minimale est fournie dans l'option suivante). Ensuite, on cherche la meilleure tâche qui respecte ce minimum.
- **`min_lexico`** : valeur minimale des préférences individuelles à atteindre. Si 0, le logiciel calculera automatiquement la valeur maximale du minimum possible des préférences individuelles et l'utilisera comme seuil. Attention : laisser 0 si `probleme_exact` est désactivé (sauf si vous savez ce que vous faites).

Augmenter la pondération du minimum :

- **`min_prof_enabled`** : `True`/`False`. Si `True`, compte plusieurs fois la préférence minimale parmi l'ensemble des professeurs dans la fonction objectif (le nombre de fois est défini par `nbr_min_count`). En donnant plus de poids au minimum, l'idée est d'encourager la maximisation de la pire des préférences. **Important** : ne peut pas être activé si lexicographic_optimization est activé.
- **`nbr_min_count`** : nombre entier strictement positif. Le nombre de fois que le minimum est ajouté à la fonction objectif si `min_prof_enabled` est `True`. **Remarque** : il n'existe pas de meilleure valeur pour ce paramètre. Elle dépend des paramètres de la tâche courante et ne peut être décidée que de façon empirique.

#### Autres paramètres

- **`allow_negative_preferences`** : `True`/`False`. Si `False`, aucun professeur ne peut se voir attribuer un cours pour lequel il a indiqué une préférence négative. Attention : désactiver cette option peut rendre la génération de la tâche impossible. L'activer rend généralement la résolution plus rapide.
- **`enforce_positive_with_negative`** : `True`/`False`. Si `True`, garantit que, si un professeur se voit attribuer un cours pour lequel il a donné une préférence négative, il dispose également d'une préférence positive de valeur au moins équivalente pour compenser (évite d'attribuer uniquement des cours avec des préférences négatives).
- **`max_hours`** : nombre entier. Limite supérieure du nombre d'heures attribuées à chaque professeur (p. ex. `14`).

#### Génération de tâches alternatives (`taches_alternatives`)

- **`nbr_taches_alternatives`** : nombre entier positif ou nul. Nombre de solutions alternatives à générer (p. ex. `3`). Si 0, aucune solution alternative n'est générée.
- **`diversite`** : nombre entier strictement positif (p. ex. `1`). Un facteur plus élevé produit, en théorie, des tâches alternatives plus variées. Un facteur trop élevé peut rendre la résolution difficile voir impossible.

**Remarque** : les tâches alternatives peuvent être plus longues à générer. Si l'optimisation lexicographique est activée, il est même possible que des tâches alternatives ne puissent pas être générées (s'il n'en existe aucun garantissant une maximisation de la préférence minimale).

### Paramètres du solveur (`solver`)

- **`time_limit`** : temps limite en secondes pour le solveur (p. ex. `240`). Après ce délai, le solveur renvoie une solution potentiellement non optimale. **Remarque** : dans le cas d'une solution non optimale, il est important de vérifier manuellement la validité de la tâche produite. Le logiciel ne le fait pas pour l'instant.
- **`gap`** : tolérance du solveur afin d'estimer l'optimalité d'une solution (p. ex. `0.001`). Une tolérance élevée produit des solutions plus rapidement, mais qui ne sont pas nécessairement réellement optimales. Inversement, une petite tolérance garantit l'optimalité du résultat, mais peut augmenter drastiquement le temps de résolution.  
- **`afficher_log`** : `True`/`False`. Si `True`, affiche les logs du solveur dans le terminal.
- **`sauvegarder_log`** : `True`/`False`. Si `True`, sauvegarde les logs du solveur dans un fichier.
- **`fichier_log`** : nom du fichier de log si `sauvegarder_log` est `True` (p. ex. `log.txt`).

### Exemple d'un fichier `config.yaml` valide

```

#################################################
#
# Générateur de tâches
# Fichier de configuration
#
#################################################


#################################################
# Fichiers d'entrée/sortie
#################################################

files:
  input: "tache.xls" # Avec l'extension
  output: "resultat_tache" # Sans extension, sera ajoutée automatiquement


#################################################
# Paramètres d'optimisation
#################################################

optimization:
  #################################################
  # Fonction objectif
  #################################################

  # Résoudre le problème exact, où la moyenne pondérée par le nombre d'heures enseignées est maximisée
  # Potentiellement lent à résoudre (penser à ajuster le temps limite dans la section "solver" ci-dessous)
  probleme_exact: True

  # Tenter de prendre en compte les libérations dans le problème approché (ne peut pas être activé si probleme_exact est activé). Rend la résolution potentiellement plus lente. Utile seulement si les libérés ont l'air lésés dans la solution approchée
  liberation_enabled: False


  #################################################
  # Maximisation du minimum des préférences individuelles
  #################################################

  # Optimisation lexicographique : maximiser la préférence minimale en premier, puis la moyenne pondérée des préférences
  # Attention : si `probleme_exact` est désactivé, le minimum optimal peut ne pas être le véritable minimum
  lexicographic_optimization: False
  min_lexico: 0 # Valeur minimale des préférences individuelles à atteindre. Si 0, le logiciel calculera automatiquement la valeur maximale du minimum possible des préférences individuelles et l'utilisera comme seuil. Attention : laisser 0 si `probleme_exact` est désactivé (sauf si vous savez ce que vous faites)

  # Alternative (incompatible avec lexicographic_optimization) : tenter de maximiser le minimum des préférences en même temps que la moyenne
  min_prof_enabled: False # Ne peut pas être activé si lexicographic_optimization est activé
  nbr_min_count: 50 # Nombre de fois que le minimum est compté dans la fonction objectif


  #################################################
  # Autres paramètres
  #################################################

  # Autoriser l'attribution de cours pour lesquels une préférence négative a été donnée
  allow_negative_preferences: False

  # Si un professeur a une préférence négative pour un cours attribué, s'assurer qu'il dispose également d'une préférence positive d'une valeur au moins équivalente
  enforce_positive_with_negative: False
  
  # Nombre maximum d'heures par professeur
  max_hours: 15


#################################################
# Génération de tâches alternatives
#################################################

taches_alternatives: # Paramètres pour la génération de tâches alternatives après la première optimisation

  # Nombre de solutions alternatives à générer
  nbr_taches_alternatives: 2
  # Facteur de diversité
  diversite: 1 # Nombre entier >= 1. Un facteur plus élevé produit, en théorie, des tâches alternatives plus variées. Un facteur trop élevé peut rendre la résolution difficile voir impossible


#################################################
# Paramètres du solveur
#################################################

solver:
  # Limite de temps pour chaque optimisation (en secondes)
  time_limit: 1000

  # Tolérance
  gap: 0.001

  # Log
  afficher_log: True
  sauvegarder_log: False
  fichier_log: "log.txt"

```
