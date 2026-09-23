## Comment ça fonctionne ?

Le générateur utilise le solveur HiGHS (via la bibliothèque Python `pulp`) pour produire des tâches. HiGHS est un logiciel libre conçu pour résoudre des modèles de programmation linéaire et de programmation en nombres entiers mixtes. En effet, le problème de la tâche peut être modélisé comme un problème de programmation en nombres entiers mixtes en définissant judicieusement la fonction objectif.

### Rappel sur la signification des préférences (dans le logiciel de la tâche d'Éric et le nouveau)

Si \( p \) désigne la préférence exprimée par le professeur \( i \) pour le cours \( j \) dans le fichier de préférences, on définit la préférence pénalisée au carré `pref(i,j)` (aussi appelée ppc) par :
$$
ppc = pref(i,j) = 1 - \frac{(2 - p)^2}{12} - \frac{(2 - p)}{6}
$$

Le tableau ci‑dessous donne la correspondance entre la préférence exprimée (entre -2 et 2) et la préférence pénalisée au carré utilisée dans l'optimisation :

| Préférence exprimée | -2 | -1   | 0   | 1   | 2 |
|---------------------|----|------|-----|-----|---|
| ppc                 | -1 | -1/4 | 1/3 | 3/4 | 1 |

**Remarque :** l'idée derrière la préférence pénalisée au carré est de pénaliser plus fortement les préférences négatives, afin d'encourager l'attribution de cours pour lesquels les préférences exprimées sont positives ou nulles lors de l'optimisation. Le logiciel tend ainsi à attribuer aux professeurs des cours qu'ils veulent vraiment.

Pour un professeur \( i \), la préférence affichée dans les scénarios de tâches correspond à la moyenne pondérée par le nombre d'heures enseignées des `pref(i,j)` :
$$
\frac{1}{\text{nbr d'heures enseignées par le prof \( i \)}} \sum_{\text{cours } j} \text{nbr d'heures du cours \( j \) enseignées} \times pref(i,j)
$$

### Qu'essaie de faire ce logiciel

Un problème de programmation linéaire consiste à optimiser une fonction objectif **linéaire** sous certaines contraintes.

Dans le cas du problème de la tâche, il s'agit alors de **maximiser** la moyenne pondérée des préférences des professeurs du département, qui est une fonction linéaire. Il ne reste qu'à définir les contraintes (intervalles de CI, nombre de préparations, etc.) et à laisser `HiGHS` résoudre le problème.

L'avantage majeur de cette approche est sa rapidité (la solution optimale étant atteinte presque instantanément). De plus, si le problème n'est pas résoluble sous les contraintes imposées, `HiGHS` permet de le savoir immédiatement.


**Remarque :** il est très simple de modifier la fonction objectif pour optimiser selon d'autres critères. Le logiciel offre par exemple de maximiser la préférence minimale parmi l'ensemble des professeurs ou encore le nombre de professeurs totalement satisfaits.