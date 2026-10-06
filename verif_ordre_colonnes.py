"""
verif_ordre_colonnes.py — Vérifie que le traitement est robuste à un
CHANGEMENT D'ORDRE des colonnes du classeur.

Ce script construit un classeur de départ minimal où chaque onglet de secteur
présente un ordre de colonnes DIFFÉRENT (et des libellés synonymes), puis
applique les règles 1 à 6. Chaque application doit être exclue par la
règle attendue, et l'onglet d'exclusion doit rester cohérent.

Verrou également la migration du 2026-02-10 : « FortiClient » et
« FortiMonitor Agent » doivent relever de la RÈGLE 3 (catégorie C) et non
de la règle 1 ; « FortiToken » doit rester conservé (garde-fou).

À lancer après toute modification de colonnes.py, d'une règle, ou du
traitement : c'est le seul garde-fou contre une régression vers un accès
aux colonnes par POSITION FIXE (row[0], row[1], ...), qui produirait des
résultats faux — voire une perte de données — de façon silencieuse.

Le vrai classeur n'est JAMAIS touché : tout se déroule dans un dossier
temporaire, supprimé à la fin. Code de retour 0 = succès, 1 = échec.

Usage : py verif_ordre_colonnes.py
"""

import shutil
import subprocess
import tempfile
from pathlib import Path

import colonnes as col

BASE_DIR = Path(__file__).resolve().parent

# --- Ordres de colonnes volontairement différents et libellés synonymes ---
# Rappel : l'ordre canonique du classeur réel est
# Logiciel | Manufacturier | Ordinateur | Lieu | Utilisateur | Department
ORDRE_REEL = ["Logiciel", "Manufacturier", "Ordinateur", "Lieu",
              "Utilisateur", "Department"]

# Les onglets ont chacun un ordre permutationné différent, et certains
# utilisent des synonymes (Nom d'hôte, Fabricant, Application, Département).
VARIANTES = [
    ["Logiciel", "Manufacturier", "Ordinateur", "Lieu", "Utilisateur",
     "Department"],
    ["Department", "Utilisateur", "Lieu", "Ordinateur", "Manufacturier",
     "Logiciel"],
    ["Application", "Fabricant", "Nom d'hôte", "Lieu", "Utilisateur",
     "Département"],
    ["Lieu", "Logiciel", "Utilisateur", "Department", "Ordinateur",
     "Manufacturier"],
]

# Lignes sources, toujours exprimées dans l'ordre réel.
LIGNES = {
    "Wahta": [
        ("7-Zip 23.01", "Igor Pavlov", "PC-WAHTA-01", "École Wahta'",
         "jean.dupont", "TI"),
        ("Microsoft 365 Apps for business", "Microsoft Corporation",
         "PC-WAHTA-01", "École Wahta'", "jean.dupont", "TI"),
        ("Autodesk AutoCAD 2024", "Autodesk", "PC-WAHTA-02",
         "École Wahta'", "marc.troti", "TI"),
        ("VLC media player", "VideoLAN", "PC-WAHTA-02", "École Wahta'",
         "marc.troti", "TI"),
        ("Adobe Acrobat Reader DC", "Adobe Inc.", "PC-WAHTA-02",
         "École Wahta'", "marc.troti", "TI"),
        # Doublon exact sur le même hôte : doit finir dans « Excluded Regle05 ».
        ("7-Zip 23.01", "Igor Pavlov", "PC-WAHTA-01", "École Wahta'",
         "jean.dupont", "TI"),
    ],
    "TI": [
        ("FortiClient VPN", "Fortinet", "PC-TI-01", "Bureau TI",
         "alice.tremblay", "TI"),
        # FortiMonitor : previously in Regle01, moved to Regle03 (category C).
        # A regression test: if this ever goes back to Regle01, the removal in
        # _EXTRAS_SUBSTRINGS (Regle 1) has been undone.
        ("FortiMonitor Agent", "Fortinet", "PC-TI-01", "Bureau TI",
         "alice.tremblay", "TI"),
        # FortiToken is deliberately NOT targeted by any rule (garde-fou).
        ("FortiToken", "Fortinet", "PC-TI-01", "Bureau TI",
         "alice.tremblay", "TI"),
        ("Microsoft Teams", "Microsoft Corporation", "PC-TI-02", "Bureau TI",
         "bob.gagnon", "TI"),
        ("Notepad++", "Notepad++ Team", "PC-TI-02", "Bureau TI",
         "bob.gagnon", "TI"),
        # La règle 6 exige des libellés EXACTS (motifs ancrés) : on utilise
        # donc des noms réellement présents dans ses listes de motifs.
        ("Python 3.12 Executables", "Python Software Foundation",
         "PC-TI-02", "Bureau TI", "bob.gagnon", "TI"),
        ("Microsoft Windows Feature Experience Pack", "Microsoft Corporation",
         "PC-TI-02", "Bureau TI", "bob.gagnon", "TI"),
    ],
}

# logiciel -> onglet d'exclusion attendu (None = doit être conservée).
# NOTA : « FortiClient » et « FortiMonitor Agent » relèvent de la RÈGLE 3
# (catégorie C) depuis le 2026-02-10 : leurs motifs ont été retirés de la
# règle 1, qui décrit les composantes livrées avec Windows 11. « FortiToken »
# reste volontairement conservé (non listé dans la règle 3).
ATTENDU = {
    "7-Zip 23.01": "Excluded Regle03",
    "VLC media player": "Excluded Regle03",
    "Adobe Acrobat Reader DC": "Excluded Regle03",
    "Microsoft 365 Apps for business": "Excluded Regle02",
    "Microsoft Teams": "Excluded Regle02",
    "Microsoft Windows Feature Experience Pack": "Excluded Regle01",
    "FortiClient VPN": "Excluded Regle03",
    "FortiMonitor Agent": "Excluded Regle03",
    "Python 3.12 Executables": "Excluded Regle06",
    "Notepad++": None,        # application d'affaires : conservée
    "FortiToken": None,       # garde-fou : non ciblé par la règle 3
}

SCRIPTS = ("colonnes.py", "appliquer_regle01.py", "appliquer_regle02.py",
           "appliquer_regle03.py", "appliquer_regle04.py",
           "appliquer_regle05.py", "appliquer_regle06.py")

LEURRE = "Excluded Regle99"  # onglet ignoré par toutes les règles


def _reordonner(ligne, entete):
    """Réordonne une ligne (ordre réel) vers l'ordre de `entete`."""
    base = dict(zip(ORDRE_REEL, ligne))
    return [base[libelle] for libelle in entete]


def construire_classeur(chemin):
    """Crée un classeur multi-onglets aux ordres de colonnes variés."""
    import openpyxl

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for i, (secteur, lignes) in enumerate(LIGNES.items()):
        entete = VARIANTES[i % len(VARIANTES)]
        ws = wb.create_sheet(secteur)
        ws.append(entete)
        for ligne in lignes:
            ws.append(_reordonner(ligne, entete))

    # Onglet leurre : ignoré par les règles (préfixe « Excluded »).
    leurre = wb.create_sheet(LEURRE)
    leurre.append(ORDRE_REEL + ["Raison"])
    leurre.append(["LEURRE"] * 6 + ["Ne doit jamais être traité"])

    # Dernier onglet de secteur : le traitement doit s'y arrêter.
    nd = wb.create_sheet("Non défini")
    nd.append(VARIANTES[1])
    for ligne in LIGNES["Wahta"]:
        nd.append(_reordonner(ligne, VARIANTES[1]))
    wb.save(chemin)


def verifier_entetes(chemin):
    """Retourne les en-têtes de « Wahta » et « TI » (contrôle du test)."""
    import openpyxl

    wb = openpyxl.load_workbook(chemin, read_only=True)
    try:
        return [
            tuple(next(wb[nom].iter_rows(min_row=1, max_row=1,
                                         values_only=True)))
            for nom in ("Wahta", "TI")
        ]
    finally:
        wb.close()


def executer_regles(dossier):
    """Applique les règles 1 à 6 dans le dossier de travail."""
    for nom in SCRIPTS[1:]:
        resultat = subprocess.run(
            ["py", str(dossier / nom)],
            cwd=dossier, capture_output=True, text=True,
            encoding="utf-8", errors="replace",
        )
        if resultat.returncode != 0:
            raise AssertionError(
                f"ÉCHEC de {nom} (code {resultat.returncode}) :\n"
                f"{resultat.stdout[-2000:]}\n{resultat.stderr[-1500:]}"
            )


def verifier_resultats(chemin):
    """Contrôle l'exclusion de chaque application et l'en-tête des exclusions.

    Retourne (exclus, echecs).
    """
    import openpyxl

    wb = openpyxl.load_workbook(chemin, read_only=True)
    try:
        exclus = {}   # {onglet d'exclusion: {logiciels}}
        entetes = {}  # {onglet d'exclusion: en-tête}
        for nom in wb.sheetnames:
            if not nom.startswith("Excluded"):
                continue
            lignes = wb[nom].iter_rows(values_only=True)
            entete = next(lignes, None)
            if entete is None:
                continue
            entetes[nom] = [str(h).strip() for h in entete]
            # Résolution par libellé : c'est précisément ce que l'on teste.
            idx_log = col.index_colonnes(entete, ("logiciel",),
                                         onglet=nom)["logiciel"]
            ensemble = set()
            for row in lignes:
                v = col.valeur(row, idx_log)
                if v:
                    ensemble.add(str(v).strip())
            exclus[nom] = ensemble

        echecs = []

        # 1. Chaque application est arrivée dans le bon onglet d'exclusion.
        for logiciel, onglet_attendu in ATTENDU.items():
            trouve = next((n for n, libs in exclus.items()
                           if logiciel in libs), None)
            if trouve != onglet_attendu:
                echecs.append(f"  « {logiciel} » : attendu {onglet_attendu}, "
                              f"trouvé {trouve}")

        # 2. Le leurre ne doit avoir été ni modifié ni infiltré : sa ligne
        #    reste seule dans SON onglet et ne doit apparaître dans aucun
        #    autre (ce qui prouverait qu'une règle l'aurait traité).
        for nom, libs in exclus.items():
            if "LEURRE" in libs and nom != LEURRE:
                echecs.append(f"  Ligne leurre retrouvée dans {nom} "
                              f"(une règle a traité l'onglet {LEURRE} !)")
        if exclus.get(LEURRE) != {"LEURRE"}:
            echecs.append(f"  Onglet {LEURRE} modifié : "
                          f"{sorted(exclus.get(LEURRE, set()))}")

        # 3. Tous les onglets d'exclusion partagent le même en-tête.
        attendu_entete = ORDRE_REEL + ["Raison"]
        for nom, entete in entetes.items():
            if nom == LEURRE:
                continue  # leurre, laissé volontairement tel quel
            if entete != attendu_entete:
                echecs.append(f"  {nom} : en-tête {entete} != {attendu_entete}")

        # 4. Aucune perte : 9 lignes distinctes doivent avoir été exclues
        #    (les lignes de « Wahta » sont traitées deux fois : dans l'onglet
        #     « Wahta » puis dans « Non défini »).
        total_exclues = sum(len(v) for n, v in exclus.items() if n != LEURRE)
        if total_exclues != 9:
            echecs.append(f"  Total de lignes exclues = {total_exclues} "
                          f"(attendu 9)")

        return exclus, echecs
    finally:
        wb.close()


def main():
    echecs = []
    dossier = Path(tempfile.mkdtemp(prefix="verif_ordre_col_"))
    try:
        for nom in SCRIPTS:
            shutil.copy2(BASE_DIR / nom, dossier / nom)

        classeur = dossier / "User_Workstation_App_Associations_Consolidated.xlsx"
        construire_classeur(classeur)

        ordres = verifier_entetes(classeur)
        print(f"Ordre des colonnes de « Wahta » : {list(ordres[0])}")
        print(f"Ordre des colonnes de « TI »    : {list(ordres[1])}")
        if ordres[0] == ordres[1]:
            echecs.append("Les deux onglets ont le même ordre : test invalide")

        print("\nApplication des règles 1 à 6...")
        executer_regles(dossier)
        print("Terminé.\n")

        exclus, echecs_resultat = verifier_resultats(classeur)
        echecs.extend(echecs_resultat)

        print("Contenu des onglets d'exclusion :")
        for nom in sorted(exclus):
            libs = ", ".join(sorted(exclus[nom])) or "(vide)"
            print(f"  {nom} : {libs}")
    finally:
        shutil.rmtree(dossier, ignore_errors=True)

    print()
    if echecs:
        print("ÉCHEC DU TEST :")
        for ligne in echecs:
            print(ligne)
        return 1
    print("SUCCÈS — traitement correct malgré des ordres de colonnes "
          "différents et des libellés synonymes.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())