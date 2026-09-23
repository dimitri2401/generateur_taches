## Paramètres de configuration du générateur de tâches

### Configuration :

Depuis la version 0.3, la configuration peut se faire via l'interface graphique. Les options disponibles sont détaillées ci-dessous.

Il est également possible de les modifier directement dans le fichier `config.yaml`, qui peut être ouvert avec n'importe quel éditeur de texte. Dans ce cas, les informations sur les différents paramètres sont détaillées dans le fichier yaml.

**Important :** ne supprimer aucune entrée du fichier de configuration.

### Fichiers

- **`Fichier d'entrée`** : nom du fichier contenant les paramètres de la tâche, avec son extension (p. ex. `tache.xls`). C'est le fichier Excel lu par le générateur, contenant les informations sur les professeurs, les cours, les préférences, les nombres de groupes maximum, etc.
- **`Fichier de sortie`** : nom du fichier de sortie, sans extension (p. ex. `resultat_tache`). Si le fichier existe déjà dans le répertoire (p. ex. en raison d'une exécution précédente du logiciel), le nouveau fichier sera renommé en lui ajoutant un numéro.

### Pré-optimisations (optionnelles)

- **`Étapes 1 et 2`** : permet de choisir les pré-optimisations effectuées et leur ordre.
- **`Préférence minimum`** : active une pré-optimisation qui cherche, parmi toutes les tâches possibles, celle dans laquelle la préférence minimale parmi l'ensemble des professeurs est maximale (sauf si une valeur minimale est fournie dans l'option suivante). Aucune tâche générée par la suite ne pourra avoir une préférence minimale inférieure à la valeur trouvée ou fournie (avec une tolérance de 1%).
- **`Nombre de totalement satisfaits`** : active une pré-optimisation qui cherche, parmi toutes les tâches possibles, celle dans laquelle le nombre de professeurs totalement satisfaits est maximal (sauf si une valeur minimale est fournie dans l'option suivante). Aucune tâche générée par la suite ne pourra avoir un nombre de professeurs totalement satisfaits inférieur à la valeur trouvée ou fournie.
- **`Seuil min. des préférences individuelles`** : nombre réel entre -1 et 1 (p. ex. `0.65`). Si l'option `Préférence minimum` est activée, on peut fixer à la main la préférence minimale autorisée. Inscrire 0 pour laisser le logiciel trouver sa valeur optimale.
- **`Seuil min. du nbr. de profs totalement satisfaits`** : nombre entier positif ou nul (p. ex. `13`). Si l'option `Nombre de totalement satisfaits` est activée, on peut fixer à la main le nombre de professeurs totalement satisfaits minimal. Inscrire 0 pour laisser le logiciel trouver sa valeur optimale.

**Remarque** : l'ordre des pré-optimisations est important. Changer l'ordre change les tâches produites.

### Paramètres d'optimisation

- **`Exclure les préférences négatives`** : si activé, aucun professeur ne peut se voir attribuer un cours pour lequel il a indiqué une préférence négative. Attention : activer cette option peut rendre la génération de la tâche impossible.
- **`Nombre maximum d'heures par professeur`** : nombre entier. Limite supérieure du nombre d'heures attribuées à chaque professeur (p. ex. `14`).

### Tâches alternatives

- **`Nombre de tâches alternatives`** : nombre entier positif ou nul. Nombre de solutions alternatives à générer (p. ex. `3`). Si 0, aucune solution alternative n'est générée.
- **`Facteur de diversité`** : nombre entier strictement positif (p. ex. `1`). Nombre minimum de professeurs qui doivent abandonner au moins un cours qu'ils enseignaient, et ce par rapport à chacune des tâches précédemment générées. Un facteur trop élevé peut rendre la génération de tâches alternatives difficile, voire impossible.

### Solveur

- **`Sauvegarder le log`** : sauvegarde les messages du programme dans un fichier. Attention, le log précédent du même nom sera écrasé.
- **`Fichier log`** : nom du fichier de log si sa sauvegarde est activée (p. ex. `log.txt`).

### Génération

- Le bouton `Lancer` démarre la génération de tâches.
- Le bouton `Arrêter` arrête une génération en cours.

### Configuration

- Le bouton `Sauvegarder` permet de sauvegarder la configuration courante.
- Le bouton `Restaurer` permet de restaurer la configuration par défaut.