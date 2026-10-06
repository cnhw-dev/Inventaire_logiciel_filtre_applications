"""
deplacer_par_secteur.py — Deux étapes de rangement du classeur
« User_Workstation_App_Associations_Consolidated.xlsx » :

Étape 1 — Déplace chaque ligne de l'onglet source « Logiciels_affaire* »
vers un onglet de « secteur », selon la valeur de la colonne « Lieu ».
Regroupement (validé avec le répondant TI) : 18 onglets au total —
17 secteurs + « Non défini » pour les lignes dont « Lieu » est vide
ou « \\N ». Voir la table SECTEURS ci-dessous. L'étape est sautée si
l'onglet source est absent (c.-à-d. déjà exécutée).

Étape 2 — Reclasse les lignes de l'onglet « Non défini » dont la
colonne « Department » est renseignée vers l'onglet du secteur
correspondant. Par défaut, un « Department » portant le nom d'un
secteur va dans l'onglet de ce secteur; les exceptions figurent dans
la table DEPARTEMENTS (p. ex. « Archives » -> « Juridique »). Les
lignes sans « Department » (vide ou « \\N ») demeurent dans
« Non défini ».

Aucune perte de données :
- toutes les lignes sont déplacées, sans transformation;
- toute valeur « Lieu » non cartographiée est signalée et envoyée dans
  un onglet portant son nom (tronqué à 31 caractères, limite Excel);
- toute valeur « Department » non cartographiée est signalée et bloque
  l'exécution (à ajouter à la table DEPARTEMENTS avant de relancer);
- le total des lignes distribuées est vérifié avant chaque sauvegarde;
- l'exécution de l'étape 1 est refusée si un onglet cible existe déjà
  (pas de risque d'écrasement ou de duplication en cas de réexécution);
- une copie de sauvegarde horodatée du classeur est créée avant
  toute exécution réelle (comme les scripts appliquer_regle0X.py);
- l'onglet « Non défini » est positionné en dernier dans le classeur
  (les scripts appliquer_regle0X.py traitent les onglets de secteurs
  jusqu'à « Non défini » inclusivement).

Comportement (calqué sur appliquer_regle05.py) :
1. Une première passe en lecture seule (read_only) classe les lignes
   et produit le tableau de distribution (utilisé tel quel en dry-run).
2. En exécution réelle (défaut), le classeur est rouvert en écriture et
   chaque étape applicable est exécutée : étape 1, les onglets cibles
   sont créés avec la ligne d'en-tête, les lignes y sont copiées, puis
   l'onglet source est supprimé; étape 2, les lignes de « Non défini »
   sont ajoutées aux onglets secteurs existants puis « Non défini »
   est réécrit avec les seules lignes conservées. Chaque étape n'est
   sauvegardée que si le total déplacé est conforme.

Usage : py deplacer_par_secteur.py             # exécution réelle (défaut)
        py deplacer_par_secteur.py --dry-run   # simulation (aucune écriture)
"""

import argparse
import re
import shutil
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

from colonnes import index_colonnes

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK = BASE_DIR / "User_Workstation_App_Associations_Consolidated.xlsx"
SOURCE_PREFIX = "Logiciels_affaire"
NON_DEFINI = "Non défini"

# Regroupement final validé : valeur « Lieu » -> onglet secteur.
SECTEURS = {
    "Wahta": [
        "École Wahta'",
        "Ecole Wahta",
        "Wendake > École Wahta'",
        "Classe 3 anne, École Wahta",
        "Classe 3 année, École Whata",
        "Classe 5 année, École Whata",
    ],
    "TI": [
        "Bureau TI",
        "Bureau TI, Bâtiment administratif CNW",
        "Wendake > Centre Administratif - CNW > Magasin TI",
    ],
    "STI": ["Bureau  STI"],
    "Santé et mieux-être": [
        "Bureaux DSME, Bâtiment  Santé Wendake",
        "Bureau 1109, Bâtiment Santé Wendake",
        "Maison des Jeûnes , DSME",
    ],
    "CHSLD": ["CHSLD Wendake, DSME"],
    "RMS": ["RMS", "RMS , DSME"],
    "CDFM": ["CDFM"],
    "Onyionhwentsiio": ["Bureau Nionwentsïo"],
    "Finance": ["Bureau Finance", "Bureau Finances"],
    "Juridique": ["Bureau Juridique"],
    "Politique": ["Bureau politique"],
    "RH": ["Bureau RH"],
    "Centre Admin": ["Wendake > Centre Administratif - CNW"],
    "PDM": ["Bureau PDM"],
    "Développement économique": ["Station EKO , Wendake"],
    "Agoshin": ["Comptoir Agoshin"],
    "Communications": ["Bureau Communication"],
}

# Table de correspondance normalisée (espaces réduits, casse ignorée).
CORRESPONDANCE = {}
for _secteur, _valeurs in SECTEURS.items():
    for _valeur in _valeurs:
        CORRESPONDANCE[" ".join(_valeur.split()).casefold()] = _secteur

# Étape 2 — valeur « Department » -> onglet secteur : exceptions à la
# règle par défaut (un « Department » portant le nom d'un secteur va
# dans l'onglet de ce secteur). Validé avec le demandeur :
DEPARTEMENTS = {
    "Archives": "Juridique",
}

# Table de correspondance normalisée « Department » -> secteur :
# chaque secteur est sa propre cible, plus les exceptions ci-dessus.
CORRESPONDANCE_DEPT = {
    " ".join(_secteur.split()).casefold(): _secteur for _secteur in SECTEURS
}
for _dept, _secteur in DEPARTEMENTS.items():
    CORRESPONDANCE_DEPT[" ".join(_dept.split()).casefold()] = _secteur

# Ordre d'affichage / de création des onglets : secteurs puis Non défini.
ORDRE_ONGLETS = list(SECTEURS) + [NON_DEFINI]

INVALID_SHEET_CHARS = re.compile(r'[:\\/?*\[\]]')


def _norm(value):
    """Normalise une valeur : espaces réduits, casse ignorée."""
    return " ".join(("" if value is None else str(value)).split()).casefold()


def _est_non_defini(norm):
    """Vrai si la valeur normalisée est vide ou « \\N »."""
    return norm in ("", "\\n")


def _onglet_secours(valeur, existants):
    """Nom d'onglet de secours (<= 31 car., sans collision) pour une
    valeur « Lieu » non cartographiée — garantit l'absence de perte."""
    base = INVALID_SHEET_CHARS.sub("", str(valeur)).strip()[:31] or "Valeur inconnue"
    nom, i = base, 2
    while nom in existants:
        suffixe = f" ({i})"
        nom = base[: 31 - len(suffixe)] + suffixe
        i += 1
    existants.add(nom)
    return nom


def analyser():
    """Passe en lecture seule : classe les lignes par onglet cible.

    Retourne (nom_source, en-tête, index Lieu, comptes, inconnues,
    lignes_vides), ou None si l'onglet source est absent (étape 1 déjà
    exécutée). Lève une exception si la colonne « Lieu » est
    introuvable.
    """
    wb = openpyxl.load_workbook(WORKBOOK, read_only=True)
    try:
        source = next(
            (ws for ws in wb.worksheets if ws.title.startswith(SOURCE_PREFIX)),
            None,
        )
        if source is None:
            return None
        lignes = source.iter_rows(values_only=True)
        entete = next(lignes)
        # Résolution par libellé d'en-tête : l'ordre des colonnes peut varier
        # d'une extraction à l'autre (voir colonnes.py).
        idx_lieu = index_colonnes(entete, ("lieu",), onglet=source.title)["lieu"]
        comptes, inconnues = Counter(), Counter()
        lignes_vides = 0
        for row in lignes:
            if all(c is None for c in row):
                lignes_vides += 1
                continue
            norm = _norm(row[idx_lieu])
            if _est_non_defini(norm):
                comptes[NON_DEFINI] += 1
            else:
                cible = CORRESPONDANCE.get(norm)
                if cible is None:
                    inconnues[row[idx_lieu]] += 1
                    cible = f"?? {row[idx_lieu]}"
                comptes[cible] += 1
        return source.title, entete, idx_lieu, comptes, inconnues, lignes_vides
    finally:
        wb.close()


def afficher_distribution(comptes, inconnues, lignes_vides):
    """Affiche le tableau de distribution et retourne le total des lignes."""
    total = sum(comptes.values())
    print("\nDistribution des lignes par onglet cible :")
    print("-" * 55)
    for nom in ORDRE_ONGLETS:
        ligne = f"  {nom:<25} {comptes.get(nom, 0):>10,} lignes"
        print(ligne.replace(",", " "))
    for nom in comptes:
        if nom.startswith("?? "):
            ligne = f"  {nom:<25} {comptes[nom]:>10,} lignes  <- NON CARTOGRAPHIE"
            print(ligne.replace(",", " "))
    print("-" * 55)
    print(f"  {'TOTAL':<25} {total:>10,} lignes".replace(",", " "))
    if lignes_vides:
        print(f"\n  ({lignes_vides} ligne(s) entièrement vide(s) ignorée(s))")
    if inconnues:
        print("\n  AVERTISSEMENT — valeurs « Lieu » non cartographiées :")
        for valeur, n in sorted(inconnues.items(), key=lambda kv: str(kv[0])):
            print(f"    {valeur!r} : {n} ligne(s)")
        print("  Elles seront placées dans des onglets de secours (aucune perte).")
    else:
        print("\n  Toutes les valeurs « Lieu » sont cartographiées. Aucune perte.")
    return total


def appliquer(nom_source, entete, idx_lieu, total_attendu):
    """Exécution réelle : crée les onglets cibles, y déplace les lignes,
    supprime l'onglet source et sauvegarde après vérification du total."""
    wb = openpyxl.load_workbook(WORKBOOK)
    try:
        existantes = [n for n in ORDRE_ONGLETS if n in wb.sheetnames]
        if existantes:
            raise RuntimeError(
                "Onglet(s) cible(s) déjà présent(s), exécution refusée "
                f"(risque de duplication) : {existantes}"
            )
        source = wb[nom_source]
        onglets = {}
        for nom in ORDRE_ONGLETS:
            ws = wb.create_sheet(nom)
            ws.append(list(entete))
            onglets[nom] = ws
        secours = {}
        noms_existants = set(wb.sheetnames)
        deplaces, lues = Counter(), 0
        debut = time.time()
        for row in source.iter_rows(min_row=2, values_only=True):
            if all(c is None for c in row):
                continue
            lues += 1
            norm = _norm(row[idx_lieu])
            if _est_non_defini(norm):
                cible = NON_DEFINI
            else:
                cible = CORRESPONDANCE.get(norm)
                if cible is None:
                    if norm not in secours:
                        nom = _onglet_secours(row[idx_lieu], noms_existants)
                        ws = wb.create_sheet(nom)
                        ws.append(list(entete))
                        onglets[nom] = ws
                        secours[norm] = nom
                        print(f"  Onglet de secours créé : « {nom} »")
                    cible = secours[norm]
            onglets[cible].append(list(row))
            deplaces[cible] += 1
            if lues % 20000 == 0:
                print(f"  ... {lues} lignes traitées ({time.time() - debut:.0f} s)")
        total_deplace = sum(deplaces.values())
        if total_deplace != total_attendu or lues != total_attendu:
            raise RuntimeError(
                "Vérification échouée : "
                f"{total_deplace} déplacées / {lues} lues / {total_attendu} "
                "attendues. Classeur NON sauvegardé."
            )
        wb.remove(source)
        # Garantir que « Non défini » est le dernier onglet du classeur :
        # les scripts appliquer_regle0X.py traitent les onglets de
        # secteurs jusqu'à « Non défini » inclusivement.
        idx = wb.sheetnames.index(NON_DEFINI)
        wb.move_sheet(NON_DEFINI, offset=len(wb.sheetnames) - 1 - idx)
        wb.save(WORKBOOK)
        print(f"\n  Vérification OK : {total_deplace} lignes déplacées, aucune perte.")
        print(f"  Onglet source « {nom_source} » supprimé. Classeur sauvegardé.")
        print(f"  « {NON_DEFINI} » positionné en dernier onglet du classeur.")
    finally:
        wb.close()


 


def analyser_departements():
    """Passe en lecture seule (étape 2) : classe les lignes de l'onglet
    « Non défini » selon la colonne « Department ».

    Retourne (index Department, comptes, inconnus, total, restants), ou
    None si l'onglet « Non défini » est absent. Lève une exception si la
    colonne « Department » est introuvable.
    """
    wb = openpyxl.load_workbook(WORKBOOK, read_only=True)
    try:
        if NON_DEFINI not in wb.sheetnames:
            return None
        ws = wb[NON_DEFINI]
        lignes = ws.iter_rows(values_only=True)
        entete = next(lignes)
        # Résolution par libellé d'en-tête : l'ordre des colonnes peut varier
        # d'une extraction à l'autre (voir colonnes.py).
        idx_dept = index_colonnes(
            entete, ("department",), onglet=NON_DEFINI)["department"]
        comptes, inconnus = Counter(), Counter()
        restants = 0
        for row in lignes:
            if all(c is None for c in row):
                continue
            norm = _norm(row[idx_dept])
            if _est_non_defini(norm):
                restants += 1
                continue
            cible = CORRESPONDANCE_DEPT.get(norm)
            if cible is None:
                inconnus[row[idx_dept]] += 1
                cible = f"?? {row[idx_dept]}"
            comptes[cible] += 1
        total = sum(comptes.values()) + restants
        return idx_dept, comptes, inconnus, total, restants
    finally:
        wb.close()


def afficher_distribution_departements(comptes, inconnus, total, restants):
    """Affiche le tableau de reclassement de l'étape 2 (par Department)."""
    print(f"\nReclassement de « {NON_DEFINI} » selon « Department » :")
    print("-" * 55)
    for nom in ORDRE_ONGLETS:
        if comptes.get(nom):
            ligne = f"  {nom:<25} {comptes[nom]:>10,} lignes"
            print(ligne.replace(",", " "))
    for nom in comptes:
        if nom.startswith("?? "):
            ligne = f"  {nom:<25} {comptes[nom]:>10,} lignes  <- NON CARTOGRAPHIE"
            print(ligne.replace(",", " "))
    ligne = f"  {'(demeure dans Non défini)':<25} {restants:>10,} lignes"
    print(ligne.replace(",", " "))
    print("-" * 55)
    print(f"  {'TOTAL':<25} {total:>10,} lignes".replace(",", " "))
    if inconnus:
        print("\n  AVERTISSEMENT — valeurs « Department » non cartographiées :")
        for valeur, n in sorted(inconnus.items(), key=lambda kv: str(kv[0])):
            print(f"    {valeur!r} : {n} ligne(s)")
        print("  Ajoutez-les à la table DEPARTEMENTS avant l'exécution.")
    else:
        print("\n  Toutes les valeurs « Department » sont cartographiées.")


def appliquer_departements(idx_dept, total_attendu):
    """Exécution réelle de l'étape 2 : déplace les lignes de « Non défini »
    dont le « Department » est cartographié vers l'onglet secteur
    correspondant (déjà existant), réécrit « Non défini » avec les seules
    lignes conservées, puis sauvegarde après vérification du total."""
    wb = openpyxl.load_workbook(WORKBOOK)
    try:
        nd = wb[NON_DEFINI]
        manquants = [s for s in SECTEURS if s not in wb.sheetnames]
        if manquants:
            raise RuntimeError(
                f"Onglet(s) secteur(s) absent(s) : {manquants}. "
                "Exécutez d'abord l'étape 1 (déplacement par « Lieu »)."
            )
        deplaces = Counter()
        conservees = []
        for row in nd.iter_rows(min_row=2, values_only=True):
            if all(c is None for c in row):
                continue
            norm = _norm(row[idx_dept])
            if _est_non_defini(norm):
                conservees.append(list(row))
                continue
            cible = CORRESPONDANCE_DEPT.get(norm)
            if cible is None:
                raise RuntimeError(
                    "Valeur « Department » non cartographiée : "
                    f"{row[idx_dept]!r}. Classeur NON sauvegardé."
                )
            wb[cible].append(list(row))
            deplaces[cible] += 1
        total_deplace = sum(deplaces.values())
        if total_deplace + len(conservees) != total_attendu:
            raise RuntimeError(
                "Vérification échouée : "
                f"{total_deplace} déplacées + {len(conservees)} conservées "
                f"!= {total_attendu} attendues. Classeur NON sauvegardé."
            )
        # Réécrit « Non défini » avec les seules lignes conservées.
        if nd.max_row > 1:
            nd.delete_rows(2, nd.max_row - 1)
        for row in conservees:
            nd.append(row)
        wb.save(WORKBOOK)
        print(
            f"\n  Vérification OK : {total_deplace} lignes reclassées, "
            f"{len(conservees)} conservées dans « {NON_DEFINI} »."
        )
        print("  Classeur sauvegardé.")
    finally:
        wb.close()


def main():
    parser = argparse.ArgumentParser(
        description="Déplace les lignes vers des onglets de secteurs : "
        "étape 1 selon « Lieu » (onglet source), étape 2 selon "
        "« Department » (reclassement de « Non défini »)."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Simulation : affiche la distribution sans modifier le classeur.",
    )
    args = parser.parse_args()

    print(f"Classeur : {WORKBOOK.name}")

    # Étape 1 : déplacement par « Lieu » (sautée si l'onglet source est
    # absent, c.-à-d. si l'étape a déjà été exécutée).
    resultat1 = analyser()
    if resultat1 is None:
        print(
            f"Aucun onglet « {SOURCE_PREFIX}* » : étape 1 déjà effectuée, "
            "elle est sautée."
        )
    else:
        nom_source, entete, idx_lieu, comptes, inconnues, vides = resultat1
        print(
            f"Onglet source : « {nom_source} » "
            f"(colonne « Lieu » : index {idx_lieu})"
        )
        total1 = afficher_distribution(comptes, inconnues, vides)

    # Étape 2 : reclassement de « Non défini » par « Department ».
    resultat2 = analyser_departements()
    if resultat2 is None:
        print(f"\nOnglet « {NON_DEFINI} » absent : étape 2 sans objet.")
    else:
        idx_dept, comptes2, inconnus2, total2, restants = resultat2
        afficher_distribution_departements(comptes2, inconnus2, total2, restants)

    if resultat1 is None and resultat2 is None:
        print("\nRien à faire.")
        return 0

    if args.dry_run:
        print("\nMode DRY-RUN : aucune modification apportée au classeur.")
        print("Relancer sans --dry-run pour exécuter le déplacement.")
        return 0
    if resultat1 is not None and inconnues:
        print("\nExécution refusée : valeurs « Lieu » non cartographiées.")
        print("Ajoutez-les à la table SECTEURS avant de relancer le script.")
        return 1
    if resultat2 is not None and inconnus2:
        print("\nExécution refusée : valeurs « Department » non cartographiées.")
        print("Ajoutez-les à la table DEPARTEMENTS avant de relancer le script.")
        return 1

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = WORKBOOK.with_name(f"{WORKBOOK.stem}_backup_{stamp}.xlsx")
    shutil.copy2(WORKBOOK, backup)
    print(f"\nCopie de sauvegarde : {backup.name}")

    print("\nMode APPLY : déplacement en cours...")
    if resultat1 is not None:
        appliquer(nom_source, entete, idx_lieu, total1)
    if resultat2 is not None:
        appliquer_departements(idx_dept, total2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
