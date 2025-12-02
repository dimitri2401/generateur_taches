## Paramètres de configuration du générateur de tâches

**Remarque :** les paramètres disponibles sont voués à être ajustés ou à disparaître au fil des expérimentations.

### Fichier de configuration :

Ce document décrit les paramètres disponibles dans le fichier de configuration `config.yaml` du générateur de tâches et explique leur effet. Le fichier de configuration peut être ouvert avec un éditeur de texte.

**Important :** ne supprimer aucune entrée du fichier de configuration.

### Fichiers d'entrée / sortie (`files`)
- **`input`** : nom du fichier contenant les paramètres de la tâche, avec son extension (par ex. `tache.xls`). C'est le fichier Excel lu par le générateur, contenant les informations sur les professeurs, les cours, les préférences, les nombres de groupes, etc.
- **`output`** : nom du fichier de sortie, sans extension (par ex. `resultat_tache`). Si le fichier existe déjà dans le répertoire (dû à une exécution préalable du logiciel par exemple), le logiciel renommera le nouveau fichier en lui attachant un numéro.

### Paramètres d'optimisation (`optimization`)

Les premiers paramètres permettent de définir la fonction objectif à maximiser. La fonction servant de base (mais qui n'est généralement pas le meilleur choix) est donnée par :
$$
\sum_{\text{prof } i} \sum_{\text{cours } j} \text{nbr de groupes du cours $j$ attribués} \times pref(i , j)
$$

Les différents paramètres permettent de modifier cette fonction et ces paramètres peuvent être combinés.

Ci-dessous sont listées les valeurs possibles et leur signification.

  - **`hours_enabled`** : `True`/`False`. Pondérer la préférence pour un cours par le nombre d'heures de ce cours. Chaque terme à l'intérieur de la fonction objectif est ainsi multiplié par la durée du cours. C'est généralement une bonne chose de l'activer, car cette option simule bien la moyenne des préférences exprimées dans les scénarios.
  - **`liberation_enabled`** : `True`/`False`. Essaie de simuler l'effet des libérations. Les professeurs libérés ayant moins de groupes ou d'heures, leur préférence a parfois un peu moins de poids dans le total. Chaque terme à l'intérieur de la fonction objectif est ainsi divisé par le pourcentage de libération du professeur.
  - **`min_cours_enabled`** : `True`/`False`. Ajoute la préférence minimale pour un cours attribué à la fonction objectif un certain nombre de fois (définit par `nbr_min_count`). En donnant plus de poids au minimum, l'idée est d'encourager la maximisation de la pire des préférences pour un cours attribué. **Remarque :** cette option ne peut pas être activée en même temps que la suivante.
  - **`min_prof_enabled`** : `True`/`False`. Ajoute la préférence minimale parmi l'ensemble des professeurs à la fonction objectif un certain nombre de fois (définit par `nbr_min_count`). En donnant plus de poids au minimum, l'idée est d'encourager la maximisation de la préférence la plus basse parmi les professeurs. **Remarque :** cette option ne peut pas être activée en même temps que la précédente.
  - **`nbr_min_count`** : nombre entier strictement positif. Le nombre de fois que le minimum est ajouté à la fonction objectif si `min_enabled` est `True`.

**Remarque :** il n'existe a priori pas de meilleure combinaison de ces paramètres et il faut les ajuster au fil des expérimentations.

- **`allow_negative_preferences`** : `True`/`False`. Si `False`, aucun professeur ne peut se voir attribuer un cours sur lequel il a mis une préférence négative. Attention, désactiver cette option pourrait rendre la tâche impossible à générer.

- **`enforce_positive_with_negative`** : `True`/`False`. Si `True`, s'assure que si un professeur se voit attribuer un cours avec une préférence négative, il ait au moins une préférence positive aussi élevée pour compenser (empêche d'attribuer uniquement des négatifs sans contrepartie).

- **`max_hours`** : nombre entier. Limite supérieure du nombre d'heures attribuées à chaque professeur (par ex. `14`).

##### Conseils pour les paramètres d'optimisation
- Essayer les différents paramètres d'optimisations. Il est difficile de prévoir lequel sera le plus efficace a priori. Certains paramètres peuvent toutefois allonger la durée de la résolution.
- Essayer de générer des tâches sans préférences négatives `allow_negative_preferences = False`. Si elles s'avèrent décevantes, autoriser les négatifs. Dans ce cas, essayer aussi d'activer `enforce_positive_with_negative`.

### Génération de tâches alternatives (`taches_alternatives`)
- **`epsilon_strategy`** : `True`/`False`. Si `True`, le générateur génère `nbr_taches_alternatives` tâches supplémentaires presque optimales. Pour ce faire, il relance l'optimiseur en imposant la contrainte `fonction objectif ≤ fonction objectif atteinte précédente - ɛ`, avec  `epsilon_value` défini ci-dessous.
- **`epsilon_value`** : nombre réel positif (par ex. `0.0001`). Valeur du epsilon utilisée pour réduire le maximum de la fonction objectif.
- **`nbr_taches_alternatives`** : nombre entier strictement positif. Nombre de solutions alternatives à générer (par ex. `3`).

##### Usage typique
- Désactiver dans un premier temps les tâches alternatives, donc `epsilon_strategy = False` et changer les paramètres du fichier de tâche jusqu'à obtenir une tâche satisfaisante.
- Ensuite seulement, générer des alternatives quasi-optimales.
- Remarque : les tâches alternatives sont généralement plus longues à générer. 

### Paramètres du solveur (`solver`)
- **`time_limit`** : temps limite en secondes pour le solveur (pr ex. `240`). Après ce temps, le solveur retourne une solution (potentiellement très) non-optimale.
- **`afficher_log`** : `True`/`False`. Afficher les logs du solveur dans le terminal.
- **`sauvegarder_log`** : `True`/`False`. Sauvegarder les logs du solveur dans un fichier.
- **`fichier_log`** : nom du fichier log si `sauvegarder_log` est `True` (par ex. `log.txt`).

##### Recommandations pour les paramètres du solveur
- Si les tâches sont trop longues à produire, augmenter le `time_limit` ou relâcher certaines contraintes (dans le fichier de paramètres ou, par exemple, en activant `allow_negative_preferences`).
- Si le logiciel est exécuté dans un terminal, il est généralement informatif d'activer `afficher_log`, mais peu utile d'activer `sauvegarder_log`.

### Exemple d'un fichier `config.yaml` valide

```

# Configuration pour le générateur de tâches

# Fichiers d'entrée/sortie
files:
  input: "tache.xls"
  output: "resultat_tache" # Sans extension, sera ajoutée automatiquement

# Paramètres d'optimisation
optimization:
  # Fonction objectif - Facteurs à inclure
  # Pondérer la préférence pour un cours par le nombre d'heures de ce cours
  hours_enabled: True

  # Tenter de prendre en compte les libérations (plus utile si hours_enabled est True)
  liberation_enabled: False

  # Tenter de maximiser également le minimum des préférences
  # Attention : ne pas activer les deux options en même temps
  min_cours_enabled: True # Cherche à augmenter la pire des préférences parmi l'ensemble des cours attribués
  min_prof_enabled: False # Chercher à augmenter la pire des préférences parmi l'ensemble des professeurs
  
  # Si l'une des deux options précédentes est activée, nombre de fois que le minimum est compté dans la fonction objectif.
  nbr_min_count: 10

  # Autoriser l'attribution de cours pour lesquels une préférence négative a été donnée.
  allow_negative_preferences: False

  # Si un professeur a une préférence négative pour un cours attribué, s'assurer qu'il dispose également d'une préférence positive de valeur au moins équivalente.
  enforce_positive_with_negative: False
  
  # Nombre d'heures maximum par professeur
  max_hours: 14

# Génération de tâches alternatives
taches_alternatives:
  # Utiliser un epsilon pour générer des tâches alternatives
  epsilon_strategy: True
  epsilon_value: 0.0001

  # Nombre de solutions alternatives à générer
  nbr_taches_alternatives: 2

# Paramètres du solveur
solver:
  # Limite de temps en secondes
  time_limit: 240

  # Log
  afficher_log: True
  sauvegarder_log: False
  fichier_log: "log.txt"

```
