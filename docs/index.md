## Générateur de tâches

### Informations générales

Ce logiciel propose une approche alternative à celle utilisée habituellement au département de mathématiques du Collège Montmorency (aussi connue sous le nom de « le logiciel d'Éric »). Il repose sur des principes d'optimisation linéaire, dans le but d'obtenir plus efficacement des tâches.

Cette approche offre certains avantages :

- **Rapidité :** la solution optimale est généralement atteinte en quelques minutes, voire quelques secondes pour le problème approché et des tâches faciles. Plusieurs solutions alternatives peuvent ensuite être produites rapidement.
- **Faisabilité :** si les contraintes rendent la tâche impossible, le logiciel le signale presque instantanément.
- **Facilité d'itération :** l'itération est simplifiée, permettant de tester rapidement l'effet des différents paramètres de la tâche qui peuvent être modifiés.
- **Flexibilité :** il est simple de définir de nouvelles contraintes (par exemple, n'avoir aucun prof à -1).

Bien évidemment, cette approche est nouvelle et n'a été testée que sur un petit nombre de tâches (contrairement à l'ancien logiciel qui aura bientôt 20 ans). Des problèmes inattendus subsistent certainement encore.

L'idée n'est donc pas de remplacer l'ancienne façon de procéder, mais de l'accompagner.

### À venir

- Une nouvelle version du solveur HiGHS devrait offrir la parallélisation du problème sous peu. De gros gains de performance sont attendus.
- Un module afin de vérifier la validité des tâches produites reste à programmer.

### Licence

Le code source de ce logiciel est distribué sous licence **GPLv3**. Vous êtes donc encouragé à le modifier, à condition de respecter les termes de la licence. Plus d'informations sur la page [À propos](about.md).