import sys
import subprocess
import threading
import queue
import os
import signal
import PySimpleGUI as sg
import webbrowser  # Pour ouvrir la documentation dans le navigateur
from ruamel.yaml import YAML  # On utilise ruamel.yaml pour préserver les commentaires


# -----------------------------------------------
# Éviter que l'interface ne soit floue sous Windows
# -----------------------------------------------

if sys.platform == "win32":
    import ctypes

    ctypes.windll.shcore.SetProcessDpiAwareness(1)


# -----------------------------------------------
# Définitions utiles
# -----------------------------------------------

# Fichiers de configuration
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    BUNDLE_DIR = sys._MEIPASS
    USER_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    USER_DIR = BUNDLE_DIR

CONFIG_FILE = os.path.join(USER_DIR, "config.yaml")  # Fichier utilisateur
CONFIG_DEFAULT_FILE = os.path.join(BUNDLE_DIR, "config_default.yaml")  # Configuration par défaut

# Options possibles pour les étapes de pré-optimisation
STEP_OPTIONS = ["prof_min", "nbr_satisfaits"]

# Libellés des étapes de pré-optimisation
STEP_LABELS = {
    "prof_min": "Préférence minimum",
    "nbr_satisfaits": "Nombre de totalement satisfaits",
}
STEP_KEYS = {v: k for k, v in STEP_LABELS.items()}  # Associe les libellés aux clés correspondantes dans le YAML


# -----------------------------------------------
# Fichier YAML
# -----------------------------------------------

ryaml = YAML()
ryaml.preserve_quotes = True
ryaml.boolean_representation = ["False", "True"]


def load_config():  # Charge le fichier YAML
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return ryaml.load(f)


# -----------------------------------------------
# Layout
# -----------------------------------------------


def build_layout(cfg):
    # On lit les valeurs du fichier YAML
    fic = cfg["fichiers"]
    opt = cfg["optimisation"]
    ta = cfg["taches_alternatives"]
    log = cfg["log"]

    ordre = list(opt["preoptimisations"])
    step1_val = ordre[0] if ordre else "Aucune"
    step1_label = STEP_LABELS.get(step1_val, "Aucune")
    if len(ordre) >= 2:
        step2_label = STEP_LABELS.get(ordre[1], "Aucune")
    else:
        step2_label = "Aucune"
    step2_options = ["Aucune"] + [STEP_LABELS[s] for s in STEP_OPTIONS if s != step1_val]

    left_col = sg.Column(
        [
            [
                sg.Frame(
                    "Fichiers",
                    [
                        [sg.Sizer(0, 5)],
                        [sg.Text("Fichier d'entrée :", size=(15, 1)), sg.Input(fic["entree"], key="-ENTREE-", size=(30, 1), tooltip="Nom du fichier xls contenant les paramètres (avec son extension)."), sg.FileBrowse("Parcourir…", file_types=(("Fichiers Excel", "*.xls"), ("Tous", "*.*")), target="-ENTREE-", size=(10, 1))],
                        [sg.Text("Fichier de sortie :", size=(15, 1)), sg.Input(fic["sortie"], key="-SORTIE-", size=(30, 1), tooltip="Nom du fichier de sortie, sans extension (elle sera ajoutée automatiquement).")],
                        [sg.Sizer(0, 10)],
                    ],
                    pad=(10, 10),
                    expand_x=True,
                )
            ],
            [
                sg.Frame(
                    "Pré-optimisations",
                    [
                        [
                            [sg.Sizer(0, 5)],
                            sg.Text("Étape 1 :", size=(10, 1)),
                            sg.Combo(
                                ["Aucune"] + list(STEP_LABELS.values()),
                                default_value=step1_label,
                                key="-STEP1-",
                                size=(35, 1),
                                auto_size_text=False,
                                readonly=True,
                                enable_events=True,
                                tooltip=("Première optimisation avant de maximiser la moyenne.\n  Préférence minimum : maximise la préférence minimale parmi tous les professeurs.\n  Nombre totalement satisfaits : maximise le nombre de professeurs totalement satisfaits."),
                            ),
                        ],
                        [
                            sg.Text("Étape 2 :", size=(10, 1)),
                            sg.Combo(
                                step2_options,
                                default_value=step2_label,
                                key="-STEP2-",
                                size=(35, 1),
                                auto_size_text=False,
                                readonly=True,
                                enable_events=True,
                                disabled=(step1_val == "Aucune"),
                                tooltip=("Deuxième optimisation, appliquée après l'étape 1.\nDisponible uniquement si une optimisation a été choisie à l'étape 1."),
                            ),
                        ],
                        [sg.Sizer(0, 3)],
                        [sg.Text("Seuil min. des préférences individuelles :", size=(40, 1)), sg.Input(str(opt["seuil_prof_min"]), key="-SEUIL_PROF_MIN-", size=(8, 1), disabled=("prof_min" not in ordre), tooltip=("Préférence minimale garantie à chaque professeur.\nSi 0, la valeur maximale atteignable est calculée automatiquement."))],
                        [sg.Text("Seuil min. du nbr. de profs totalement satisfaits :", size=(40, 1)), sg.Input(str(opt["seuil_nbr_satisfaits"]), key="-SEUIL_NBR_SATISFAITS-", size=(8, 1), disabled=("nbr_satisfaits" not in ordre), tooltip=("Nombre minimal de professeurs devant être totalement satisfaits.\nSi 0, le maximum atteignable est calculé automatiquement."))],
                        [sg.Sizer(0, 10)],
                    ],
                    pad=(10, 10),
                    expand_x=True,
                )
            ],
            [
                sg.Frame(
                    "Paramètres d'optimisation",
                    [
                        [sg.Sizer(0, 10)],
                        [sg.Checkbox("Exclure les préférences négatives", default=bool(opt["exclure_preferences_negatives"]), key="-PREF_NEG-", tooltip="Si activé, aucun cours avec une préférence strictement négative ne peut être attribué.", pad=((0, 5), (0, 0)))],
                        [sg.Sizer(0, 3)],
                        [sg.Text("Nombre maximum d'heures par professeur :", size=(38, 1)), sg.Input(str(opt["nbr_heures_max"]), key="-NBR_HEURES_MAX-", size=(8, 1), tooltip="Charge horaire maximale pouvant être attribuée à un professeur.")],
                        [sg.Sizer(0, 10)],
                    ],
                    pad=(10, 10),
                    expand_x=True,
                )
            ],
            [
                sg.Frame(
                    "Tâches alternatives",
                    [
                        [sg.Sizer(0, 10)],
                        [sg.Text("Nombre de tâches alternatives :", size=(38, 1)), sg.Input(str(ta["nbr_taches_alternatives"]), key="-NBR_ALT-", size=(8, 1), tooltip="Nombre de tâches supplémentaires à générer.")],
                        [sg.Text("Facteur de diversité :", size=(38, 1)), sg.Input(str(ta["diversite"]), key="-DIVERSITE-", size=(8, 1), tooltip=("Entier >= 1. Nombre minimum de professeurs qui doivent abandonner au moins un cours par rapport aux tâches précédemment générées."))],
                        [sg.Sizer(0, 10)],
                    ],
                    pad=(10, 10),
                    expand_x=True,
                )
            ],
            [
                sg.Frame(
                    "Log",
                    [
                        [sg.Sizer(0, 10)],
                        [sg.Checkbox("Sauvegarder le log dans un fichier", default=bool(log["sauvegarder_log"]), key="-SAUVEGARDER_LOG-", enable_events=True, tooltip="Enregistre les messages du programme dans un fichier texte. Attention, le log précédent du même nom sera écrasé.", pad=((0, 5), (0, 0)))],
                        [sg.Sizer(0, 3)],
                        [sg.Text("Fichier log :", size=(38, 1)), sg.Input(log["fichier_log"], key="-FICHIER_LOG-", size=(20, 1), disabled=not bool(log["sauvegarder_log"]), tooltip="Nom du fichier dans lequel les messages sont sauvegardés.")],
                        [sg.Sizer(0, 10)],
                    ],
                    pad=(10, 10),
                    expand_x=True,
                )
            ],
            [
                sg.Frame(
                    "Génération",
                    [
                        [sg.Sizer(0, 5)],
                        [sg.Button("Lancer", key="-RUN-", size=(10, 1), pad=(10, 10), button_color=("white", "#2E7D32"), tooltip="Lancer l'optimisation."), sg.Button("Arrêter", key="-KILL-", size=(10, 1), disabled=True, pad=(10, 10), button_color=("white", "#C62828"), tooltip="Arrêter l'optimisation.")],
                        [sg.Sizer(0, 5)],
                    ],
                    pad=(10, 10),
                    expand_x=True,
                ),
                sg.Frame(
                    "Configuration",
                    [
                        [sg.Sizer(0, 5)],
                        [sg.Button("Sauvegarder", key="-SAVE-", size=(10, 1), pad=(10, 10), tooltip="Sauvegarder les paramètres"), sg.Button("Restaurer", key="-RESTORE-", size=(12, 1), pad=(10, 10), tooltip="Restaurer les paramètres par défaut.\nAttention, cela écrasera le fichier config.yaml actuel.")],
                        [sg.Sizer(0, 5)],
                    ],
                    pad=(10, 10),
                    expand_x=True,
                ),
            ],
            [sg.Button("Cliquer ici pour consulter la documentation.", key="-LINK-", pad=(10, 10), tooltip="La documentation s'ouvrira dans votre fureteur.")],
        ],
        vertical_alignment="top",
        expand_y=True,
    )

    right_col = sg.Column(
        [
            [
                sg.Multiline(  # Terminal
                    "",
                    key="-OUTPUT-",
                    autoscroll=True,
                    disabled=True,
                    font=("Courier New", 9),
                    background_color="black",
                    text_color="white",
                    size=(100, None),
                    expand_x=True,
                    expand_y=True,
                )
            ],
        ],
        expand_x=True,
        expand_y=True,
    )

    layout = [[left_col, right_col]]
    return layout


# -----------------------------------------------
# Sauvegarde de la configuration
# -----------------------------------------------


def save_config(values, cfg):
    try:
        seuil_prof_min = float(values["-SEUIL_PROF_MIN-"])
        seuil_nbr_satisfaits = int(values["-SEUIL_NBR_SATISFAITS-"])
        nbr_heures_max = int(values["-NBR_HEURES_MAX-"])
        nbr_taches_alternatives = int(values["-NBR_ALT-"])
        diversite = int(values["-DIVERSITE-"])
    except (TypeError, ValueError):
        raise ValueError("Les paramètres numériques entrés contiennent des valeurs invalides.")

    if seuil_nbr_satisfaits < 0 or nbr_heures_max <= 0 or nbr_taches_alternatives < 0 or diversite < 1:
        raise ValueError("Les paramètres numériques entrés contiennent des valeurs invalides.")

    cfg["fichiers"]["entree"] = values["-ENTREE-"]
    cfg["fichiers"]["sortie"] = values["-SORTIE-"]

    step1 = STEP_KEYS.get(values["-STEP1-"], "Aucune")
    step2 = STEP_KEYS.get(values["-STEP2-"], "Aucune")
    ordre = []
    if step1 != "Aucune":
        ordre.append(step1)
        if step2 != "Aucune":
            ordre.append(step2)
    preopti = cfg["optimisation"].get("preoptimisations")
    if preopti is None:
        cfg["optimisation"]["preoptimisations"] = ordre
    else:
        preopti[:] = ordre

    cfg["optimisation"]["seuil_prof_min"] = seuil_prof_min
    cfg["optimisation"]["seuil_nbr_satisfaits"] = seuil_nbr_satisfaits
    cfg["optimisation"]["exclure_preferences_negatives"] = values["-PREF_NEG-"]
    cfg["optimisation"]["nbr_heures_max"] = nbr_heures_max

    cfg["taches_alternatives"]["nbr_taches_alternatives"] = nbr_taches_alternatives
    cfg["taches_alternatives"]["diversite"] = diversite

    cfg["log"]["sauvegarder_log"] = values["-SAUVEGARDER_LOG-"]
    cfg["log"]["fichier_log"] = values["-FICHIER_LOG-"]

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        ryaml.dump(cfg, f)


# -----------------------------------------------
# Exécution de main.py et sauvegarde de sa sortie
# -----------------------------------------------


def kill_main(proc):  # Termine le processus
    if proc is None or proc.poll() is not None:
        return
    if sys.platform == "win32":
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    else:
        os.killpg(os.getpgid(proc.pid), signal.SIGKILL)


def main_py_worker(output_queue):  # Lance main.py comme un sous-processus et enregistre sa sortie dans output_queue
    if getattr(sys, "frozen", False):
        cmd = [sys.executable, "--nogui"]  # L'option --nogui lance main.py directement
    else:
        cmd = [sys.executable, os.path.join(BUNDLE_DIR, "main.py")]
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        cwd=USER_DIR,
        bufsize=1,
        text=True,
        start_new_session=True,
    )

    def reader():
        for line in proc.stdout:
            output_queue.put(line)
        proc.wait()
        output_queue.put(proc.returncode if proc.returncode is not None else 0)

    threading.Thread(target=reader, daemon=True).start()
    return proc


# -----------------------------------------------
# Boucle principale
# -----------------------------------------------


def main():
    # Thème
    sg.theme("DarkGrey13")

    # On s'assure que config.yaml existe, sinon on le crée
    if not os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_DEFAULT_FILE, "r", encoding="utf-8") as src, open(CONFIG_FILE, "w", encoding="utf-8") as dst:
                dst.write(src.read())
        except Exception:
            sg.popup_error("Erreur lors de la création du fichier de configuration.")
            raise SystemExit(1)

    # On charge la configuration
    cfg = load_config()

    # On construit la fenêtre
    window = sg.Window("Générateur de tâches", build_layout(cfg), resizable=True, finalize=True)

    # Variables pour gérer main.py et sa sortie
    worker = None
    output_queue = queue.Queue()

    while True:
        event, values = window.read(timeout=33)  # Refresh toutes les 33 ms afin d'avoir un affichage fluide du terminal

        try:
            while True:
                line = output_queue.get_nowait()
                if isinstance(line, int):  # Si c'est un entier, c'est certainement la fin du processus avec un code de retour
                    window["-OUTPUT-"].update("\n--- Processus terminé ---\n", append=True)
                    window["-RUN-"].update(disabled=False)  # On active le bouton exécuter
                    window["-KILL-"].update(disabled=True)  # On désactive le bouton arrêter
                else:
                    window["-OUTPUT-"].update(line, append=True)  # On affiche la ligne dans le terminal
        except queue.Empty:
            pass

        if event == sg.WIN_CLOSED:
            kill_main(worker)
            break

        # Préoptimisation - Étape 1 : mettre à jour le menu déroulant de l'étape 2 + activer les seuils
        elif event == "-STEP1-":
            select = values["-STEP1-"]
            if select and select != "Aucune":
                select_key = STEP_KEYS.get(select)
                other_key = [STEP_LABELS[s] for s in STEP_OPTIONS if s != select_key]
                window["-STEP2-"].update(values=["Aucune"] + other_key, value="Aucune", disabled=False)
                window["-SEUIL_PROF_MIN-"].update(disabled=(select_key != "prof_min"))
                window["-SEUIL_NBR_SATISFAITS-"].update(disabled=(select_key != "nbr_satisfaits"))
            else:
                window["-STEP2-"].update(value="Aucune", disabled=True)
                window["-SEUIL_PROF_MIN-"].update(disabled=True)
                window["-SEUIL_NBR_SATISFAITS-"].update(disabled=True)

        # Préoptimisation - Étape 2 : activer les seuils
        elif event == "-STEP2-":
            step1_key = STEP_KEYS.get(values["-STEP1-"], "Aucune")
            step2_key = STEP_KEYS.get(values["-STEP2-"], "Aucune")
            window["-SEUIL_PROF_MIN-"].update(disabled=("prof_min" not in (step1_key, step2_key)))
            window["-SEUIL_NBR_SATISFAITS-"].update(disabled=("nbr_satisfaits" not in (step1_key, step2_key)))

        # Activer/désactiver le champ fichier log
        elif event == "-SAUVEGARDER_LOG-":
            window["-FICHIER_LOG-"].update(disabled=not values["-SAUVEGARDER_LOG-"])

        # Lancer main.py
        elif event == "-RUN-":
            try:
                save_config(values, cfg)
            except ValueError as error:
                sg.popup_error(str(error), title="Configuration invalide")
                continue
            window["-OUTPUT-"].update("")
            worker = main_py_worker(output_queue)
            window["-RUN-"].update(disabled=True)
            window["-KILL-"].update(disabled=False)

        # Arrêter main.py
        elif event == "-KILL-":
            if worker and worker.poll() is None:
                kill_main(worker)
                window["-OUTPUT-"].update("\n--- Processus arrêté ---\n", append=True)
            window["-RUN-"].update(disabled=False)
            window["-KILL-"].update(disabled=True)

        # Sauvegarder la configuration
        elif event == "-SAVE-":
            try:
                save_config(values, cfg)
            except ValueError as error:
                sg.popup_error(str(error), title="Configuration invalide")
                continue
            window["-OUTPUT-"].update("\n--- Configuration sauvegardée ---\n", append=True)

        # Restaurer le fichier config.yaml par défaut
        elif event == "-RESTORE-":
            try:
                with open(CONFIG_DEFAULT_FILE, "r", encoding="utf-8") as src, open(CONFIG_FILE, "w", encoding="utf-8") as dst:
                    dst.write(src.read())
                cfg = load_config()
                # Recharger la fenêtre avec les nouvelles valeurs
                window.close()
                window = sg.Window("Générateur de tâches", build_layout(cfg), resizable=True, finalize=True)
                window["-OUTPUT-"].update("\n--- config.yaml recréé avec les valeurs par défaut ---\n", append=True)
                continue
            except Exception:
                window["-OUTPUT-"].update("\n--- Erreur lors de la restauration du fichier de configuration ---\n", append=True)

        # Ouvrir la documentation dans le fureteur
        elif event == "-LINK-":
            webbrowser.open_new_tab("https://dimitri2401.github.io/generateur_taches/")

    window.close()


if __name__ == "__main__":  # Mode nogui : si l'exécutable est lancé avec --nogui, on exécute directement main.py. Incompatible avec l'option console=False de pyinstaller sous Windows
    if "--nogui" in sys.argv:
        import main as main_module

        main_module.main()
    else:
        main()
