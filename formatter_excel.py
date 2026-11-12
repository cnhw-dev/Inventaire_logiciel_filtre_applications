"""
formatter_excel.py — Formate les onglets de secteurs du classeur
« User_Workstation_App_Associations_Consolidated.xlsx » (jusqu'à l'onglet
« Non défini » inclusivement) :

1. tri des lignes selon « Department », « Utilisateur », « Ordinateur »
   puis « Logiciel » (croissant, casse et accents ignorés, valeurs vides
   en fin — tri stable : les lignes à clés égales gardent leur ordre
   d'origine);
2. conversion de la plage en tableau Excel (filtre d'en-tête sur toutes
   les colonnes, style « TableStyleMedium9 » à bandes de lignes);
3. figeage de la ligne d'en-tête (freeze_panes = « A2 »).

Les onglets « Excluded RegleXX », « Sommaire », « Feuil1 », etc. ne sont
pas touchés : le traitement s'arrête après « Non défini ». Si « Non
défini » est absent du classeur, le formatage est refusé (la liste des
secteurs est alors incertaine).

Aucune perte de données :
- le tri ne fait que réordonner les lignes existantes, sans en modifier
  le contenu; le nombre de lignes de chaque onglet est vérifié avant la
  sauvegarde (le classeur n'est PAS sauvegardé en cas d'écart);
- le script est idempotent : les tableaux et filtres existants d'un
  onglet de secteur sont retirés avant recréation, donc une réexécution
  reformate proprement (pas d'erreur de nom de tableau dupliqué);
- un onglet sans données (en-tête seule, p. ex. « Agoshin ») n'est pas
  converti en tableau (signalé), mais sa ligne d'en-tête est figée
  comme les autres;
- un en-tête contenant une cellule vide ou un libellé dupliqué bloque
  l'exécution (les noms de colonnes d'un tableau Excel doivent être
  uniques et non vides) — signalé AVANT toute écriture.

Comportement :
1. Une passe en lecture seule (read_only) analyse chaque onglet de
   secteur : résolution des colonnes de tri par libellé (voir
   colonnes.py — jamais par position), comptage des lignes et aperçu
   des premières lignes après tri. En mode --dry-run, c'est tout.
2. En exécution réelle (défaut), le classeur est ouvert une seule fois
   en écriture : chaque onglet de secteur est trié, converti en tableau
   et figé, puis le classeur est sauvegardé une seule fois à la fin.

Aucune copie de sauvegarde n'est créée ici (même politique que les
scripts appliquer_regle0X.py).

Usage : py formatter_excel.py             # exécution réelle (défaut)
        py formatter_excel.py --dry-run   # simulation (aucune écriture)
"""

import argparse
import re
import sys
import time
import unicodedata
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableStyleInfo

from colonnes import index_colonnes, valeur

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK = BASE_DIR / "User_Workstation_App_Associations_Consolidated.xlsx"
LAST_SECTOR = "Non défini"  # dernier onglet de secteur à traiter (inclusivement)
IGNORED = {"logiciels", "sommaire", "feuil1", "top 10",
           "excluded regle01", "excluded regle02", "excluded regle03",
           "excluded regle04", "excluded regle05", "excluded regle06"}

# Rôles de colonnes (voir colonnes.py) définissant l'ordre de tri :
# « Department », puis « Utilisateur », « Ordinateur » et « Logiciel ».
ROLES_TRI = ("department", "utilisateur", "hote", "logiciel")

# Style appliqué aux tableaux créés (« Blanc, Style de tableau clair 8 »).
TABLE_STYLE = "TableStyleLight8"

APERCU_LIGNES = 3  # nombre de lignes d'aperçu affichées par onglet

_INVALID_TABLE_CHARS = re.compile(r"[^0-9A-Za-z_]")


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


def _secteurs(sheetnames):
    """Retourne les onglets de secteurs jusqu'à « Non défini » inclusivement."""
    ignored = {s.casefold() for s in IGNORED}
    secteurs = []
    for name in sheetnames:
        if name.casefold() not in ignored:
            secteurs.append(name)
        if name.casefold() == LAST_SECTOR.casefold():
            break
    return secteurs


def _norm_tri(value):
    """Normalise une valeur pour la clé de tri.

    Accents retirés, espaces réduits, casse ignorée — proche du tri
    « français » d'Excel, mais déterministe. Retourne "" si vide.
    """
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(text.split()).casefold()


def _cle_ligne(row, indices):
    """Clé de tri d'une ligne : les valeurs vides passent à la fin."""
    cle = []
    for index in indices:
        norm = _norm_tri(valeur(row, index))
        cle.append((1, "") if norm == "" else (0, norm))
    return tuple(cle)


def _nom_table(nom_onglet, existants):
    """Nom de tableau unique et ASCII, dérivé du nom de l'onglet.

    Les noms de tableaux sont globaux au classeur et ne peuvent contenir
    ni espaces ni accents : on préfixe par « T_ » et on suffixe en cas de
    collision (p. ex. « T_Finance_2 »).
    """
    text = unicodedata.normalize("NFKD", nom_onglet)
    text = "".join(c for c in text if not unicodedata.combining(c))
    base = _INVALID_TABLE_CHARS.sub("_", text).strip("_") or "Onglet"
    nom, i = f"T_{base}", 2
    while nom.casefold() in existants:
        nom = f"T_{base}_{i}"
        i += 1
    existants.add(nom.casefold())
    return nom


def _verifier_entete(entete, onglet):
    """Refuse les en-têtes avec cellule vide ou libellés dupliqués.

    Les noms de colonnes d'un tableau Excel doivent être uniques (casse
    ignorée) et non vides; mieux vaut bloquer AVANT toute écriture que
    produire un classeur qu'Excel signalerait comme corrompu.
    """
    vus = {}
    for i, libelle in enumerate(entete, start=1):
        if libelle is None or str(libelle).strip() == "":
            raise RuntimeError(
                f"En-tête vide à la colonne {i} de l'onglet « {onglet} » : "
                "un tableau Excel exige un libellé par colonne. Classeur "
                "NON modifié."
            )
        cle = " ".join(str(libelle).split()).casefold()
        if cle in vus:
            raise RuntimeError(
                f"Libellé d'en-tête dupliqué « {libelle} » (colonnes "
                f"{vus[cle]} et {i}) dans l'onglet « {onglet} » : un "
                "tableau Excel exige des libellés uniques. Classeur NON "
                "modifié."
            )
        vus[cle] = i


def _abr(value, limite=45):
    """Abrège une valeur pour l'aperçu (« (vide) » si absente)."""
    if value is None:
        return "(vide)"
    text = " ".join(str(value).split())
    if len(text) > limite:
        return text[: limite - 1] + "…"
    return text



def analyser():
    """Passe en lecture seule : pour chaque onglet de secteur, résout les
    colonnes de tri, vérifie l'en-tête, compte les lignes et calcule
    l'aperçu trié.

    Retourne la liste des analyses (dicts). Lève une RuntimeError dès
    qu'une colonne de tri est introuvable/ambiguë, qu'un en-tête est
    invalide ou que « Non défini » est absent — avant toute écriture.
    """
    wb = openpyxl.load_workbook(WORKBOOK, read_only=True)
    try:
        secteurs = _secteurs(wb.sheetnames)
        if not secteurs or secteurs[-1].casefold() != LAST_SECTOR.casefold():
            raise RuntimeError(
                f"L'onglet « {LAST_SECTOR} » est introuvable : la liste "
                "des secteurs est incertaine, formatage refusé. Classeur "
                "NON modifié."
            )
        analyses = []
        for nom in secteurs:
            ws = wb[nom]
            lignes = ws.iter_rows(values_only=True)
            entete = next(lignes, None)
            if entete is None:
                raise RuntimeError(
                    f"L'onglet « {nom} » est vide (aucune en-tête). "
                    "Classeur NON modifié."
                )
            _verifier_entete(entete, nom)
            cols = index_colonnes(entete, ROLES_TRI, onglet=nom)
            indices = tuple(cols[role] for role in ROLES_TRI)
            donnees = [row for row in lignes
                       if not all(c is None for c in row)]
            donnees.sort(key=lambda row: _cle_ligne(row, indices))
            analyses.append({
                "nom": nom,
                "nb_colonnes": len(entete),
                "indices": indices,
                "nb_lignes": len(donnees),
                "apercu": donnees[:APERCU_LIGNES],
            })
        return analyses
    finally:
        wb.close()


def appliquer(analyses):
    """Exécution réelle : tri, conversion en tableau et ligne figée pour
    chaque onglet de secteur; une seule ouverture et une seule sauvegarde.
    """
    wb = openpyxl.load_workbook(WORKBOOK)
    try:
        # Purge défensive préalable (idempotence) : retire les tableaux et
        # les filtres déjà présents sur les onglets de secteurs avant de
        # recréer un tableau par onglet.
        for info in analyses:
            ws = wb[info["nom"]]
            for nom_table in list(ws.tables):
                del ws.tables[nom_table]
            ws.auto_filter.ref = None
        # Noms de tableaux déjà utilisés ailleurs dans le classeur (les
        # noms sont globaux au classeur, pas propres à un onglet).
        existants = {t.casefold() for feuille in wb.worksheets
                     for t in feuille.tables}

        total = 0
        for i, info in enumerate(analyses, start=1):
            nom, nb_lignes = info["nom"], info["nb_lignes"]
            ws = wb[nom]
            if nb_lignes:
                lignes = [row for row in ws.iter_rows(min_row=2,
                                                      values_only=True)
                          if not all(c is None for c in row)]
                if len(lignes) != nb_lignes:
                    raise RuntimeError(
                        f"Onglet « {nom} » : {len(lignes)} ligne(s) lue(s) "
                        f"en écriture contre {nb_lignes} en analyse. "
                        "Classeur NON sauvegardé."
                    )
                lignes.sort(key=lambda row: _cle_ligne(row, info["indices"]))
                ws.delete_rows(2, ws.max_row - 1)
                for row in lignes:
                    ws.append(row)
                if ws.max_row != nb_lignes + 1:
                    raise RuntimeError(
                        f"Onglet « {nom} » : {ws.max_row - 1} ligne(s) "
                        f"réécrite(s) contre {nb_lignes} attendue(s). "
                        "Classeur NON sauvegardé."
                    )
                fin_col = get_column_letter(info["nb_colonnes"])
                ref = f"A1:{fin_col}{nb_lignes + 1}"
                table = Table(displayName=_nom_table(nom, existants),
                              ref=ref, headerRowCount=1)
                table.tableStyleInfo = TableStyleInfo(
                    name=TABLE_STYLE, showFirstColumn=False,
                    showLastColumn=False, showRowStripes=True,
                    showColumnStripes=False)
                # Filtre d'en-tête porté par le tableau (et non par la
                # feuille) : un seul autoFilter par plage, sinon Excel
                # signale un contenu corrompu à l'ouverture.
                table.autoFilter = AutoFilter(ref=ref)
                ws.add_table(table)
                print(f"  [{i}/{len(analyses)}] « {nom} » : {nb_lignes} "
                      f"ligne(s) triée(s), tableau « {table.displayName} » "
                      f"({ref}).", flush=True)
            else:
                print(f"  [{i}/{len(analyses)}] « {nom} » : aucune donnée, "
                      "tableau non créé.", flush=True)
            ws.freeze_panes = "A2"
            total += nb_lignes

        wb.save(WORKBOOK)
        print(f"\n  Vérification OK : {total} ligne(s) triée(s) sur "
              f"{len(analyses)} onglet(s), aucune perte.")
        print("  Classeur sauvegardé.")
    finally:
        wb.close()



def main():
    """Analyse les onglets de secteurs, affiche l'aperçu du tri, puis
    applique le formatage sauf en mode --dry-run."""
    parser = argparse.ArgumentParser(
        description="Formate les onglets de secteurs : tri "
        "Department/Utilisateur/Ordinateur/Logiciel, conversion en "
        "tableau avec filtre d'en-tête, ligne d'en-tête figée."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Simulation : affiche l'aperçu sans modifier le classeur.",
    )
    args = parser.parse_args()

    # Évite les UnicodeEncodeError quand la sortie est redirigée vers un
    # fichier (encodage console cp1252).
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass

    print(f"Classeur : {WORKBOOK.name}", flush=True)
    debut = time.perf_counter()
    analyses = analyser()
    print(f"Onglets de secteurs à formater (jusqu'à « {LAST_SECTOR} » "
          f"inclusivement) : {len(analyses)}")

    total = 0
    for i, info in enumerate(analyses, start=1):
        total += info["nb_lignes"]
        print(f"\n[{i}/{len(analyses)}] « {info['nom']} » — "
              f"{info['nb_lignes']} ligne(s)")
        if info["nb_lignes"] == 0:
            print("    (aucune donnée : pas de tableau, en-tête figée "
                  "seulement)")
            continue
        print("    Aperçu après tri (Department > Utilisateur > "
              "Ordinateur > Logiciel) :")
        for row in info["apercu"]:
            cellules = [_abr(valeur(row, idx)) for idx in info["indices"]]
            print("      " + " | ".join(cellules))
    print(f"\nTotal : {total} ligne(s) sur {len(analyses)} onglet(s).",
          flush=True)

    if args.dry_run:
        print("\nMode DRY-RUN : aucune modification apportée au classeur.")
        print("Relancer sans --dry-run pour appliquer le formatage.")
        return 0

    print("\nMode APPLY : formatage en cours...", flush=True)
    appliquer(analyses)
    print(f"\nTerminé en {_fmt_duree(time.perf_counter() - debut)}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

