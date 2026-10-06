"""
executer_traitement_complet.py — Exécute le traitement complet du classeur
« User_Workstation_App_Associations_Consolidated.xlsx » dans l'ordre validé :

1. déplacement des lignes vers les onglets de secteurs;
2. application des règles 1 à 6, en ordre croissant.

Le traitement s'arrête immédiatement dès qu'un script retourne un code
d'erreur. Aucun classeur de secours n'est créé ici : chaque script appelé
effectue déjà sa propre copie de sauvegarde avant son exécution réelle.

Usage : py executer_traitement_complet.py
"""

from pathlib import Path
import shutil
import subprocess
import sys
import time

BASE_DIR = Path(__file__).resolve().parent
SCRIPTS = (
    "deplacer_par_secteur.py",
    "appliquer_regle01.py",
    "appliquer_regle02.py",
    "appliquer_regle03.py",
    "appliquer_regle04.py",
    "appliquer_regle05.py",
    "appliquer_regle06.py",
)


def _fmt_duree(secondes):
    """Formate une durée en secondes sous la forme « 1 h 02 min 03 s »."""
    secondes = int(secondes)
    heures, reste = divmod(secondes, 3600)
    minutes, secondes = divmod(reste, 60)
    if heures:
        return f"{heures} h {minutes:02d} min {secondes:02d} s"
    if minutes:
        return f"{minutes} min {secondes:02d} s"
    return f"{secondes} s"


def _code_erreur(returncode):
    """Retourne un code de sortie positif utilisable par sys.exit()."""
    return returncode if returncode > 0 else 1


def executer_script(nom_script, index, total, py_launcher):
    """Exécute un script et retourne 0 en cas de succès, sinon son erreur."""
    chemin_script = BASE_DIR / nom_script
    print("\n" + "=" * 78, flush=True)
    print(f"[{index}/{total}] {nom_script}", flush=True)
    print("=" * 78, flush=True)

    debut = time.perf_counter()
    commande = [py_launcher, str(chemin_script)]
    try:
        resultat = subprocess.run(commande, cwd=BASE_DIR, check=False)
    except OSError as exc:
        print(f"\nERREUR — impossible de lancer « {nom_script} » : {exc}",
              file=sys.stderr, flush=True)
        return 1

    duree = time.perf_counter() - debut
    if resultat.returncode != 0:
        print(f"\nERREUR — « {nom_script} » a retourné le code "
              f"{resultat.returncode} après {_fmt_duree(duree)}.",
              file=sys.stderr, flush=True)
        return _code_erreur(resultat.returncode)

    print(f"\nSUCCÈS — « {nom_script} » terminé en {_fmt_duree(duree)}.",
          flush=True)
    return 0


def main():
    """Vérifie les prérequis, puis exécute tous les scripts dans l'ordre."""
    # Évite les UnicodeEncodeError quand la sortie est redirigée vers un
    # fichier (encodage console cp1252).
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass

    py_launcher = shutil.which("py")
    if py_launcher is None:
        print("ERREUR — le lanceur Windows « py » est introuvable dans le PATH.",
              file=sys.stderr)
        return 1

    manquants = [nom for nom in SCRIPTS if not (BASE_DIR / nom).is_file()]
    if manquants:
        print("ERREUR — script(s) introuvable(s) :", file=sys.stderr)
        for nom in manquants:
            print(f"  - {nom}", file=sys.stderr)
        return 1

    total = len(SCRIPTS)
    print("Traitement complet du classeur", flush=True)
    print(f"Dossier : {BASE_DIR}", flush=True)
    print(f"Lanceur : {py_launcher}", flush=True)
    print(f"Nombre d'étapes : {total}", flush=True)
    print("Les sauvegardes sont créées par chacun des scripts appelés.", flush=True)

    debut_total = time.perf_counter()
    try:
        for index, nom_script in enumerate(SCRIPTS, start=1):
            code = executer_script(nom_script, index, total, py_launcher)
            if code != 0:
                print("\n" + "!" * 78, file=sys.stderr, flush=True)
                print(f"TRAITEMENT INTERROMPU à l'étape {index}/{total} : "
                      f"{nom_script}", file=sys.stderr, flush=True)
                print("Aucun script suivant n'a été exécuté.",
                      file=sys.stderr, flush=True)
                print("!" * 78, file=sys.stderr, flush=True)
                return code
    except KeyboardInterrupt:
        print("\n\nTRAITEMENT INTERROMPU par l'utilisateur (Ctrl+C).",
              file=sys.stderr, flush=True)
        return 130

    print("\n" + "=" * 78, flush=True)
    print(f"TRAITEMENT COMPLET RÉUSSI en "
          f"{_fmt_duree(time.perf_counter() - debut_total)}.", flush=True)
    print("=" * 78, flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
