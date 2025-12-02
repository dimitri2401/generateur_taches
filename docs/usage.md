## Utilisation du générateur de tâches

### Configuration

La configuration du logiciel se fait dans le fichier `config.yaml`. Voir la page [Configuration](configuration.md) pour plus de détails.

**Important :** avant de pouvoir générer des tâches, il faut fournir au logiciel un fichier de paramètres au format `.xls`, dont le nom correspond à l'information présente dans le fichier `config.yaml` (par ex. `tache.xls`). Ce fichier doit être copié dans le répertoire contenant l'exécutable du logiciel (ou le fichier source Python).

### Lancement à partir du fichier exécutable

#### Windows

- Il est possible de simplement cliquer sur l'exécutable `generateur_taches.exe` afin de générer des tâches.
- Si le log est activé, il est toutefois recommandé d'exécuter le programme dans un terminal, afin de pouvoir surveiller le déroulement de la résolution.

**Remarque :** si Windows Defender bloque l'exécution du programme, il faut cliquer sur `Informations complémentaires` pour pouvoir autoriser le logiciel.

#### Linux

- Exécuter le programme `generateur_taches` dans un terminal.

#### MacOS

- Exécutable non testé, je n'ai pas de Mac. Il devrait toutefois fonctionner, mais il faut peut-être autoriser son exécution.

### Lancement à partir du fichier Python

- Ouvrir un terminal dans le répertoire `src`
- Exécuter `python main.py`.