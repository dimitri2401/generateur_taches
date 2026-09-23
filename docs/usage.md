## Utilisation du générateur de tâches

### Configuration

La configuration du logiciel se fait dans l'interface graphique. Voir la page [Configuration](configuration.md) pour plus de détails.

**Important :** avant de pouvoir générer des tâches, il faut fournir au logiciel un fichier de paramètres valide au format `.xls`, dont le nom correspond à celui indiqué dans l'interface graphique ou dans le fichier `config.yaml` (par ex. `tache.xls`). Ce fichier doit se trouver dans le répertoire contenant l'exécutable du logiciel (ou le fichier source Python).

### Lancement à partir du fichier exécutable

#### Windows

- Cliquer sur l'exécutable `generateur_taches.exe` afin d'ouvrir l'interface graphique.

**Remarque :** si Windows Defender bloque l'exécution du programme, il faut cliquer sur `Informations complémentaires` pour pouvoir autoriser le logiciel.

#### Linux

- Exécuter le programme `generateur_taches` dans un terminal afin d'ouvrir l'interface graphique.
- La version en ligne de commande peut être lancée depuis un terminal en utilisant l'option `--nogui`.

#### MacOS

- Exécuter le programme depuis un terminal.

### Lancement à partir du fichier Python

- Ouvrir un terminal dans le répertoire `src`.
- Exécuter `python gui.py`, ou `python main.py` pour la version en ligne de commande.

### Génération de tâches

- Une fois l'interface graphique lancée et la configuration effectuée, il suffit de cliquer sur `Lancer` pour générer des tâches.
- Le bouton `Arrêter` permet d'arrêter une génération en cours.