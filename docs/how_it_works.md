## Comment ça fonctionne ?

Le générateur utilise le solveur HiGHS (via la bibliothèque Python `pulp`) pour produire des tâches. HiGHS est un logiciel libre conçu pour résoudre des modèles de programmation linéaire et de programmation en nombres entiers mixtes.

Le problème de la tâche peut être modélisé comme un problème de programmation en nombres entiers mixtes en définissant judicieusement la fonction objectif.

### Rappel sur la signification des préférences (dans le logiciel de la tâche d'Éric et le nouveau)

Si \( p \) désigne la préférence exprimée par le professeur \( i \) pour le cours \( j \) dans le fichier de préférences, on définit la préférence pénalisée au carré `pref(i,j)` (aussi appelée ppc) par :
$$
ppc = pref(i,j) = 1 - \frac{(2 - p)^2}{12} - \frac{(2 - p)}{6}
$$

Le tableau ci‑dessous donne la correspondance entre la préférence exprimée (entre -2 et 2) et la préférence pénalisée au carré utilisée dans l'optimisation :

| Préférence exprimée | -2 | -1   | 0   | 1   | 2 |
|---------------------|----|------|-----|-----|---|
| ppc                 | -1 | -1/4 | 1/3 | 3/4 | 1 |

Remarque : l'idée derrière la préférence pénalisée au carré est de pénaliser plus fortement les préférences négatives, afin d'encourager l'attribution de cours pour lesquels les préférences exprimées sont positives ou nulles lors de l'optimisation.

Pour un professeur \( i \), la préférence personnelle affichée dans les scénarios de tâches correspond à la moyenne pondérée par le nombre d'heures enseignées des `pref(i,j)` :
$$
\frac{1}{\text{nbr d'heures enseignées par le prof \( i \)}} \sum_{\text{cours } j} \text{nbr d'heures du cours \( j \) enseignées} \times pref(i,j)
$$

### Qu'essaie de faire ce logiciel

Un problème de programmation linéaire consiste à optimiser (ici maximiser) **une seule** fonction objectif **linéaire** sous certaines contraintes. A priori, la tâche consiste à optimiser **plusieurs** fonctions (les préférences de chaque professeur) **non linéaires** (à cause de la division par le nombre d'heures totales enseignées).

Il est toutefois possible de linéariser la fonction donnant les préférences de chaque professeur, puis d'optimiser la moyenne des préférences.

Atteindre l'optimalité pour cet objectif conduit généralement à la maximisation des termes de la somme (et donc des préférences individuelles). Si ce n'est pas le cas, plusieurs options permettent d'assurer une préférence minimale satisfaisante pour chaque professeur.

**Remarques :**

- Pour accélérer le calcul, il est possible de résoudre une version approchée du problème. La résolution est alors beaucoup plus rapide et fournit des solutions essentiellement équivalentes (mais pas toujours parfaitement optimales).
- Il est très simple de modifier la fonction objectif pour optimiser selon d'autres critères.

Il reste ensuite à définir les contraintes (intervalles de CI, nombre de préparations maximum, etc.) et à laisser `HiGHS` résoudre le problème.

L'avantage majeur de cette approche est sa rapidité (quelques minutes, voire secondes pour le problème approché, pour atteindre une solution optimale). De plus, `HiGHS` permet de savoir presque instantanément si le problème est réellement solvable sous les contraintes imposées.