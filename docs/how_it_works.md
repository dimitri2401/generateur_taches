## Comment ça fonctionne ?

Le générateur utilise le logiciel `HiGHS` (via la librairie Python `pulp`) afin de produire des tâches. `HiGHS` est un logiciel libre conçu pour résoudre des modèles de programmation linéaire et de programmation mixte en nombres entiers.

Le problème de la tâche peut être approché comme un problème de programmation mixte en nombres entiers en définissant judicieusement la fonction objectif.

### Rappel sur la signification des préférences (dans le logiciel de la tâche d'Éric et le nouveau)

Si \( p \) dénote la préférence exprimée par le professeur \( i \) pour le cours \( j \) dans le fichier de préférences, on définit la préférence pénalisée au carré \( pref(i,j) \) (aussi appelée ppc) par :
$$
ppc = pref(i,j) = 1 - \frac{(2 - p)^2}{12} - \frac{(2 - p)}{6}
$$

Le tableau ci-dessous donne explicitement la correspondance entre la préférence exprimée dans le fichier (entre -2 et 2) et la préférence pénalisée au carré, utilisée dans l'optimisation (donc \( pref(i,j)  \), la ppc):

| Préférence exprimée | -2 | -1   | 0   | 1   | 2 |
|---------------------|----|------|-----|-----|---|
| ppc                 | -1 | -1/4 | 1/3 | 3/4 | 1 |

**Remarque :** l'idée derrière la préférence pénalisée au carré et pénaliser plus fortement les préférences négatives, afin de motiver l'attribution de préférences exprimées positives ou nulles lors de l'optimisation.

Pour un professeur \( i \), la préférence personnelle telle qu'affichée dans les scénarios de tâches est la moyenne, pondérée sur le nombre d'heures enseignées, des \( pref(i,j) \) :
$$
\frac{1}{\text{nbr tot d'heures enseignées par le prof $i$}} \sum_{\text{cours } j} \text{nbr d'heures du cours $j$ enseignées} \times pref(i,j)
$$

### Qu'essaie de faire ce logiciel

Un problème de programmation linéaire consiste à optimiser (ici maximiser) **une seule** fonction objectif **linéaire** sous certaines contraintes. A priori, la tâche consiste à optimiser **plusieurs** fonctions (les préférences de chaque professeur) **non linéaires** (à cause de la division par le nombre d'heures totales enseignées).

Il est toutefois possible d'approximer le problème de façon raisonnable en définissant comme fonction objectif la fonction linéaire suivante (ou une variante de cette fonction, selon les options activées):
$$
\sum_{\text{prof } i} \sum_{\text{cours } j} \text{nbr d'heures du cours $j$ enseignées} \times pref(i,j)
$$

Il s'agit donc d'optimiser la préférence globale (et pas totalement pondérée) des préférences. Atteindre l'optimalité dans cet objectif aboutit généralement à une maximisation des termes de la somme (et donc des préférences individuelles).

Une série d'options permet de changer les critères d'optimisation (tenter de maximiser également le minimum des préférences, de prendre en compte les libérations, de ne pas prendre en compte la durée des cours, etc.).

**Remarques :**

- Il est possible, mais beaucoup trop coûteux calculatoirement, de définir une fonction essentiellement égale à la véritable fonction à optimiser. Les résultats obtenus ne sont alors pas vraiment meilleurs.
- Il est très simple de changer la fonction objectif, afin de maximiser selon d'autres critères.

Il reste ensuite simplement à définir les contraintes (intervalles de CI, nombre de préparations maximum, etc) et à laisser `HiGHS` résoudre le problème.

L'avantage majeur de cette approche pour résoudre le problème de la tâche est sa rapidité (quelques secondes) avant d'atteindre une solution optimale. `HiGHS` permet de plus de savoir presque instantanément si le problème est réellement résoluble sous les contraintes imposées.