"""
appliquer_regle01.py — Applique la Règle 1 (exclusions générales : composantes
installées par défaut avec Windows 11 et mises à niveau) sur les onglets
de secteurs de « User_Workstation_App_Associations_Consolidated.xlsx »
(jusqu'à l'onglet « Non défini » inclusivement).

Comportement (conforme à Regle_01_exclusions_generales.md) :
1. Génère d'abord la liste distincte des applications correspondant à la
   règle (aucun déplacement). Cette analyse est faite en lecture seule
   (read_only) pour être rapide et économe en mémoire.
2. Traite ensuite UN secteur à la fois : le classeur est rouvert pour cet
   onglet seulement, les lignes correspondantes sont déplacées vers l'onglet
   « Excluded Regle01 » (créé une seule fois) sans transformation des données
   d'origine (une colonne finale « Raison » est ajoutée dans l'onglet
   d'exclusion), puis le classeur est sauvegardé, fermé et les ressources
   sont libérées avant de passer au secteur suivant.
3. La progression est affichée à l'écran (onglet courant, temps écoulé et
   estimation du temps restant). Aucune intervention n'est requise entre
   les secteurs.

Usage : py appliquer_regle01.py --dry-run   # simulation (aucune écriture)
        py appliquer_regle01.py             # exécution réelle
"""

import argparse
import gc
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl

from colonnes import aligner, index_colonnes, valeur

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK = BASE_DIR / "User_Workstation_App_Associations_Consolidated.xlsx"
EXCLUDED_SHEET = "Excluded Regle01"
REASON = "Règle 1 — Installé par défaut avec Windows 11 / mise à niveau Windows"
LAST_SECTOR = "Non défini"  # dernier onglet de secteur à traiter (inclusivement)
IGNORED = {"logiciels", "sommaire", "feuil1", EXCLUDED_SHEET.casefold()}


def _norm(value):
    """Normalise une valeur : casefold, apostrophes unifiées, espaces réduits."""
    text = "" if value is None else str(value)
    for ch in ("’", "‘", "`"):
        text = text.replace(ch, "'")
    return " ".join(text.split()).casefold()


# --- A) Mises à jour / mises à niveau Windows --------------------------------
_KB_RE = re.compile(r"kb\d{6,}")
_UPDATE_SUBSTRINGS = (
    "security update for microsoft",
    "cumulative update",
    "hotfix",
    "servicing stack update",
    "definition update",
    "update for microsoft",
)
_UPDATE_TOOLS_EXACT = (
    "assistant d'installation de windows 11",
    "assistant mise à jour de windows 10",
    "microsoft update health tools",
)

# --- B) Runtimes et composantes système Windows ------------------------------
_SYSTEM_SUBSTRINGS = (
    "windowsappruntime",
    "winappruntime",
    "windows app runtime",
    "microsoft ui xaml",
    "widgets platform runtime",
    "windows feature experience pack",
    "pack d'expérience de fonctionnalités windows",
    "windows web experience pack",
    "windows sdk",
    "windows software development kit",
    "microsoft edge webview",
    "microsoft webview",
    "microsoft visual c++",
    "microsoft client web",
    "microsoft crossdevice",
    "microsoft cross device",
)
_WORKLOAD_RE = re.compile(r"windows\s*workload")
_OEM_RE = re.compile(r"\boem\b")
_MS_WINDOWS_RE = re.compile(r"microsoft\.windows")

# --- C) Applications in-box Windows 11 (sous-chaînes) ------------------------
_INBOX_SUBSTRINGS = (
    "mspaint",
    "microsoft paint",
    "microsoft store",
    "storepurchaseapp",
    "microsoft people",
    "microsoft messaging",
    "microsoft pay",
    "oneconnect",
    "microsoft solitaire collection",
    "solitaire & casual games",
    "microsoft family",
    "microsoft gethelp",
    "microsoft getstarted",
    "microsoft content",
    "screensketch",
    "screen sketch",
    "stickynotes",
    "sticky notes",
    "microsoft journal",
    "microsoft pcmanager",
    "3d viewer",
    "3dviewer",
    "mixedreality",
    "microsoft todos",
    "startexperiences",
    "xbox game ui",
    "xbox tcui",
    "xbox game bar plugin",
    "xbox one smartglass",
    "game bar",
    "internet explorer",
    "windows sandbox",
    "windows subsystem for linux",
    "windows hello setup",
    "configuration de windows hello",
    "windows search",
    "windows advanced settings",
    "windows default lock screen",
    "écran de verrouillage par défaut",
    "cortana",
    "inputapp",
    "asynctextservice",
    "app resolver",
    "webprintingservice",
    "ncsiuwpapp",
    "localservicecomponents",
    "onnxruntime",
    "winmlshared",
    "engagement framework",
    "aifabric",
    "local experience pack",
    "speech pack",
    "speech translation pack",
    "ink.handwriting",
)

# --- C) Applications in-box Windows 11 (correspondance exacte) ---------------
_INBOX_EXACT = {
    "xbox",
    "narrator",
    "narrateur",
    "explorateur de fichiers",
    "file explorer",
    "paramètres",
    "settings",
    "e-mail et comptes",
    "email and accounts",
    "messagerie et comptes",
    "compte professionnel ou scolaire",
    "work or school account",
    "galerie de photos",
    "photos (preview)",
    "impression windows",
    "windows print",
    "file d'impression",
    "print queue",
    "expérience windows shell",
    "contrôle d'intégrité du pc windows",
    "contrôle parental microsoft",
    "fonctionnalités de la famille microsoft",
    "fonctionnalités famille microsoft",
    "application de verrouillage d'accès attribué",
    "application de verrouillage pour accès affecté",
    "assigned access lock app",
    "eye control",
    "contrôle visuel",
    "supprimer l'appareil en toute sécurité",
    "safely remove device",
    "captive portal flow",
    "débit du portail captif",
    "débit de la connexion réseau",
    "network connection flow",
    "examen",
    "passer un examen",
    "appeler",
    "aperçu de code-barres windows",
    "windows barcode preview",
    "credential dialog",
    "boîte de dialogue d'informations d'identification",
    "programme de résolution de l'application",
    "programme de résolution des applis",
    "pinningconfirmationdialog",
    "add folder suggestions dialog",
    "boîte de dialogue ajouter des suggestions de dossier",
    "windows",
}


# --- D) Outils, pilotes, agents et composantes diverses (ajouts) -------------
# Les motifs sont volontairement sans numéro de version afin d'exclure toutes
# les versions lorsque la version fait partie du libellé de l'application.
_EXTRAS_SUBSTRINGS = (
    # Outils éditeurs, agents et services
    "acrobat notification client",
    "adobe refresh manager",
    "antidote - connectix",            # Antidote - Connectix <version>
    "logi download assistant",
    "logitech unifying",               # Logiciel Logitech Unifying <version>
    "mozilla maintenance service",
    "autodesk cer",
    "glpi agent",                      # GLPI Agent <version>
    # NOTE : « forticlient » et « fortimonitor » ont été RETIRÉS d'ici le
    # 2026-02-10. Ces outils sont déployés par les TI, pas livrés avec
    # Windows 11 : ils relèvent de la Règle 3 (catégorie C), qui les exclut
    # désormais. Voir _FORTI_SUBSTRINGS dans appliquer_regle03.py.
    "screenconnect client",            # ScreenConnect Client (<hash>)
    "mstsc-",                          # mstsc-<GUID>
    "assistant de téléchargement",
    # Pilotes et utilitaires matériels
    "asus osd driver",
    "asuspcassistant",                 # ex. B9ECED6F.ASUSPCAssistant
    "displaylink",
    "dosebadge",
    "dts audio processing",
    "elan pointstick",
    "intelgraphicsexperience",         # ex. AppUp.IntelGraphicsExperience
    "realtek audio control",
    # Runtimes .NET / ASP.NET (toutes versions)
    "microsoft .net host",             # couvre aussi « Host FX Resolver »
    "microsoft .net runtime",
    ".net native framework package",
    ".net native runtime package",
    "asp.net core",
    "windows desktop runtime",
    # Extensions média Microsoft
    "av1videoextension",
    "avcencodervideoextension",
    "heifimageextension",
    "hevcvideoextension",
    "mpeg2videoextension",
    "rawimageextension",
    "vp9videoextensions",
    "webmediaextensions",
    "webpimageextension",
    # Composantes et applis système Microsoft / Windows
    "applicationcompatibilityenhancements",
    "bingsearch",
    "search in bing",
    "companyportal",
    "d3dmappinglayers",
    "desktopappinstaller",
    "gameassist",
    "edge update",
    "edge devtools",
    "intune management extension",
    "actionsserver",
    "sechealthui",
    "teams meeting add-in",
    "windows 11 entreprise",
    "augloop",
    "callingshellapp",
    "microsoft windows photos",
    "startmenuexperiencehost",
    "windowsalarms",
    "windowscalculator",
    "windowscamera",
    "windowscommunicationsapps",
    "windowsfeedbackhub",
    "windowsmaps",
    "windowsnotepad",
    "windowsterminal",
    "quickassist",
    "microsoftwindows.",               # ex. MicrosoftWindows.Client.CoreAI
    "module d'expérience locale",      # variante FR de « Local Experience Pack »
    "contenu microsoft",               # variante FR de « Microsoft Content »
    "click-to-run extensibility",
    "click-to-run localization",
    "officepushnotificationsutility",
    "onenote virtual printer",
    "local ai manager",
    "smartscreen",
    "roomdevice",
    "capturepicker",
    "defaultpackmsi",
    "management and security status",  # Intel(R) Management and Security Status
)

_EXTRAS_EXACT = {
    "1527c705-839a-4832-9118-54d4bd6a0c89",
    "bonjour",
    "flux du portail captif",
    "microsoft teams (personal)",
    "remote help",
    "udk package",
    "visionneuse web de l'application de bureau",
    "votre compte",
}

# « Antidote <version> - Module français » : la version est au milieu du
# libellé, d'où une expression régulière plutôt qu'une sous-chaîne.
_ANTIDOTE_MODULE_RE = re.compile(r"antidote.*module fran")


def match_rule1(logiciel, manufacturier):
    """Retourne la catégorie ('A', 'B', 'C' ou 'D') si la ligne correspond
    à la Règle 1, sinon None."""
    name = _norm(logiciel)
    if not name:
        return None
    combined = f"{name} {_norm(manufacturier)}"

    if (_KB_RE.search(name)
            or any(s in name for s in _UPDATE_SUBSTRINGS)
            or name in _UPDATE_TOOLS_EXACT):
        return "A"
    if (any(s in combined for s in _SYSTEM_SUBSTRINGS)
            or _WORKLOAD_RE.search(combined)
            or _OEM_RE.search(combined)
            or _MS_WINDOWS_RE.search(combined)):
        return "B"
    if any(s in name for s in _INBOX_SUBSTRINGS) or name in _INBOX_EXACT:
        return "C"
    if (any(s in name for s in _EXTRAS_SUBSTRINGS)
            or name in _EXTRAS_EXACT
            or _ANTIDOTE_MODULE_RE.search(name)):
        return "D"
    return None


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


def _fmt_duree(secondes):
    """Formate une durée en secondes sous la forme « 1 h 02 min 03 s »."""
    secondes = int(secondes)
    h, reste = divmod(secondes, 3600)
    m, s = divmod(reste, 60)
    if h:
        return f"{h} h {m:02d} min {s:02d} s"
    if m:
        return f"{m} min {s:02d} s"
    return f"{s} s"


def main():
    # Évite les UnicodeEncodeError quand la sortie est redirigée vers un
    # fichier (encodage console cp1252) : les caractères non représentables
    # sont remplacés plutôt que de faire planter le script.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass

    parser = argparse.ArgumentParser(
        description="Applique la Règle 1 (Windows 11 par défaut / mises à niveau).")
    parser.add_argument("--dry-run", action="store_true",
                        help="Simulation : affiche les résultats sans modifier le classeur.")
    args = parser.parse_args()

    if not args.dry_run:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = WORKBOOK.with_name(f"{WORKBOOK.stem}_backup_{stamp}.xlsx")
        shutil.copy2(WORKBOOK, backup)
        print(f"Copie de sauvegarde : {backup.name}", flush=True)

    # Étape 1 — analyse en lecture seule (rapide, peu de mémoire) :
    # liste distincte des applications correspondant à la règle et nombre
    # de lignes à déplacer par secteur. Aucun déplacement à cette étape.
    print("\nÉtape 1/2 — Analyse des onglets (lecture seule)...", flush=True)
    wb = openpyxl.load_workbook(WORKBOOK, read_only=True)
    try:
        secteurs = _secteurs(wb.sheetnames)
        print(f"Onglets de secteurs à traiter (jusqu'à « {LAST_SECTOR} ») : "
              f"{len(secteurs)}", flush=True)
        header = None
        distinct = set()
        compteurs = {}   # {onglet: nb de lignes correspondantes}
        entetes = {}     # {onglet: en-tête d'origine}
        index_par_onglet = {}  # {onglet: index des colonnes requises}
        for sheet_name in secteurs:
            ws = wb[sheet_name]
            lignes = ws.iter_rows(values_only=True)
            entete = next(lignes, None)
            if entete is not None:
                entetes[sheet_name] = list(entete)
                # Les colonnes sont localisées par leur LIBELLÉ, jamais par
                # leur position : l'ordre peut varier d'une extraction à
                # l'autre. Lève une RuntimeError explicite si une colonne
                # requise est absente.
                index_par_onglet[sheet_name] = index_colonnes(
                    entete, ("logiciel", "manufacturier"), onglet=sheet_name)
                if header is None:
                    header = list(entete)
            cols = index_par_onglet[sheet_name]
            count = 0
            for row in lignes:
                logiciel = valeur(row, cols["logiciel"])
                manufacturier = valeur(row, cols["manufacturier"])
                categorie = match_rule1(logiciel, manufacturier)
                if categorie:
                    count += 1
                    distinct.add((str(logiciel).strip(), categorie))
            compteurs[sheet_name] = count
    finally:
        wb.close()
        del wb
        gc.collect()

    print("\nListe distincte des applications correspondant à la Règle 1 :")
    for logiciel, categorie in sorted(distinct, key=lambda t: (t[1], t[0].casefold())):
        print(f"  [{categorie}] {logiciel}")
    total_planned = sum(compteurs.values())
    print(f"\nApplications distinctes : {len(distinct)}")
    print(f"Lignes à déplacer (tous secteurs) : {total_planned}", flush=True)
    if args.dry_run:
        print("Mode --dry-run : aucune modification.")
        return

    # Étape 2 — un secteur à la fois : ouverture du classeur, déplacement des
    # lignes vers « Excluded Regle01 », sauvegarde, fermeture et libération
    # des ressources avant de passer au secteur suivant (aucune pause).
    print(f"\nÉtape 2/2 — Traitement des secteurs (un à la fois, "
          f"sauvegarde après chacun) :", flush=True)
    debut = time.perf_counter()
    total_deplace = 0
    nb_secteurs = len(secteurs)
    for i, sheet_name in enumerate(secteurs, start=1):
        ecoule = time.perf_counter() - debut
        if i > 1:
            restant = ecoule / (i - 1) * (nb_secteurs - i + 1)
            progression = (f"temps écoulé {_fmt_duree(ecoule)}, "
                           f"restant estimé {_fmt_duree(restant)}")
        else:
            progression = "estimation du temps restant après le 1er onglet"
        print(f"\n[{i}/{nb_secteurs}] Onglet « {sheet_name} » ({progression})",
              flush=True)

        if compteurs[sheet_name] == 0:
            print("  -> 0 ligne à déplacer, onglet inchangé.", flush=True)
            continue

        t0 = time.perf_counter()
        wb = openpyxl.load_workbook(WORKBOOK)
        try:
            ws = wb[sheet_name]
            cols = index_par_onglet[sheet_name]
            matches = []
            for row_idx, row in enumerate(
                    ws.iter_rows(min_row=2, values_only=True), start=2):
                logiciel = valeur(row, cols["logiciel"])
                manufacturier = valeur(row, cols["manufacturier"])
                if match_rule1(logiciel, manufacturier):
                    matches.append((row_idx, row))

            excluded_ws = (wb[EXCLUDED_SHEET]
                           if EXCLUDED_SHEET in wb.sheetnames else None)
            if excluded_ws is None:
                excluded_ws = wb.create_sheet(EXCLUDED_SHEET)
                excluded_ws.append((header or []) + ["Raison"])
            for _row_idx, values in matches:
                # Réordonne la ligne vers l'en-tête de référence : les onglets
                # de secteurs peuvent présenter un ordre de colonnes différent
                # sans que l'onglet d'exclusion ne scramblie les colonnes.
                excluded_ws.append(
                    aligner(values, entetes[sheet_name], header,
                            onglet=sheet_name) + [REASON])
            # Suppression par blocs de lignes consécutives (de bas en haut) :
            # beaucoup plus rapide qu'une suppression ligne par ligne.
            indices = sorted(m[0] for m in matches)
            i = len(indices) - 1
            while i >= 0:
                j = i
                while j > 0 and indices[j - 1] == indices[j] - 1:
                    j -= 1
                ws.delete_rows(indices[j], indices[i] - indices[j] + 1)
                i = j - 1

            # Ajuste les plages des tableaux structurés et des filtres à la
            # nouvelle fin de données. Sans cela, les plages pointent au-delà
            # des lignes restantes et Excel signale un contenu corrompu à
            # l'ouverture du fichier.
            fin = max(ws.max_row, 1)
            for tbl in ws.tables.values():
                m = re.match(r"^([A-Z]+\d+):([A-Z]+)(\d+)$", tbl.ref)
                if m and int(m.group(3)) != fin:
                    nouveau = f"{m.group(1)}:{m.group(2)}{fin}"
                    tbl.ref = nouveau
                    if tbl.autoFilter is not None:
                        tbl.autoFilter.ref = nouveau
            if ws.auto_filter.ref:
                m = re.match(r"^([A-Z]+\d+):([A-Z]+)(\d+)$", ws.auto_filter.ref)
                if m and int(m.group(3)) != fin:
                    ws.auto_filter.ref = f"{m.group(1)}:{m.group(2)}{fin}"

            wb.save(WORKBOOK)
        finally:
            wb.close()
            del wb
            gc.collect()  # libère les ressources avant le secteur suivant

        total_deplace += len(matches)
        print(f"  -> {len(matches)} ligne(s) déplacée(s) et sauvegardée(s) "
              f"en {_fmt_duree(time.perf_counter() - t0)}.", flush=True)

    print(f"\nTerminé en {_fmt_duree(time.perf_counter() - debut)}.")
    print(f"Total déplacé : {total_deplace} ligne(s) -> « {EXCLUDED_SHEET} ».")
    print("Aucune perte de données : chaque ligne retirée y est copiée telle quelle.")


if __name__ == "__main__":
    main()



