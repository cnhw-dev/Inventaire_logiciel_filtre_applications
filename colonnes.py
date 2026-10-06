"""
colonnes.py — Résolution des colonnes du classeur consolidé, par libellé.

Le classeur « User_Workstation_App_Associations_Consolidated.xlsx » est
régénéré périodiquement et l'ordre de ses colonnes peut varier d'une
extraction à l'autre. Aucun script ne doit donc JAMAIS accéder à une
colonne par sa position fixe (row[0], row[1], row[2], ...) : cette
position est implicite, silencieuse, et produit un résultat FAUX — voire
une perte de données — si l'ordre change.

Ce module centralise la résolution : les colonnes sont localisées à partir
de leur libellé d'en-tête, et non de leur position.

Usage type :

    entete = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    cols = index_colonnes(entete, ("logiciel", "hote"), onglet=ws.title)
    ...
    logiciel = valeur(row, cols["logiciel"])
    hote     = valeur(row, cols["hote"])

Si une colonne requise est absente de l'en-tête, `index_colonnes` lève une
RuntimeError décrivant l'en-tête réellement lu et les libellés acceptés. Le
traitement s'interrompt AVANT toute écriture plutôt que de produire un
résultat faux — même politique que les vérifications de total déjà
présentes dans les scripts.

Trois garanties sont fournies :

1. `index_colonnes` — résolution par nom, avec synonymes (par exemple
   « Ordinateur », « Nom d'hôte » et « Computer » désignent le même rôle) ;
2. `valeur` — lecture défensive d'une ligne (retourne None si la colonne
   est absente de la ligne, au lieu de lever une IndexError) ;
3. `aligner` — réordonne une ligne vers l'ordre d'un en-tête de référence,
   afin que des onglets dont l'ordre diffère alimentent un même onglet
   d'exclusion sans y scrambler les colonnes.
"""

import unicodedata

# Libellés acceptés pour chaque rôle, après normalisation (_norm_libelle).
SYNONYMES = {
    "logiciel": (
        "logiciel", "logiciels", "nom du logiciel", "nom de l'application",
        "application", "applications", "applicatif", "applicatifs",
        "software", "produit", "produits",
    ),
    "manufacturier": (
        "manufacturier", "manufacturiers", "fabricant", "fournisseur",
        "editeur", "constructeur", "manufacturer", "vendor", "publisher",
    ),
    "hote": (
        "ordinateur", "ordinateurs", "nom d'hote", "nom de l'hote",
        "nom du poste", "nom de poste", "poste", "poste de travail",
        "computer", "hostname", "host name",
    ),
    "lieu": (
        "lieu", "emplacement", "emplacissement", "location", "site",
    ),
    "utilisateur": (
        "utilisateur", "utilisateurs", "nom d'utilisateur", "nom de l'utilisateur",
        "user", "username", "usager", "compte",
    ),
    "department": (
        "department", "departement", "service", "direction",
    ),
}


def _norm_libelle(value):
    """Normalise un libellé d'en-tête pour comparaison.

    Apostrophes typographiques unifiées, accents retirés, ponctuation
    réduite à des espaces, espaces réduits, casse ignorée. Ainsi
    « Nom d'hôte », « nom d’hote » et « NOM D'HOTE  » donnent la même clé.
    """
    text = "" if value is None else str(value)
    for ch in ("’", "‘", "`", "´"):
        text = text.replace(ch, "'")
    # NFKD décompose « é » en « e » + accent combinant ; on retire l'accent.
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = "".join(c if (c.isalnum() or c in " '") else " " for c in text)
    return " ".join(text.split()).casefold()


def libelles_attendus(role):
    """Libellés acceptés pour un rôle (pour les messages d'erreur)."""
    return SYNONYMES.get(role, (role,))


def index_colonnes(entete, roles, onglet="?"):
    """Résout les index de colonnes pour les `roles` demandés.

    `entete` : première ligne de l'onglet (itérable de libellés) ;
    `roles`  : rôles requis, par exemple ("logiciel", "manufacturier") ;
    `onglet` : nom de l'onglet, uniquement pour les messages d'erreur.

    Retourne un dict {rôle: index}. Lève une RuntimeError si un rôle
    demandé est absent de l'en-tête.
    """
    entete = list(entete or ())
    cles = [_norm_libelle(h) for h in entete]
    indices = {}
    for role in roles:
        attendus = libelles_attendus(role)
        index = next(
            (i for i, cle in enumerate(cles) if cle and cle in attendus),
            None,
        )
        if index is None:
            lus = ", ".join(repr(h) for h in entete) if entete else "(vide)"
            raise RuntimeError(
                f"Colonne « {role} » introuvable dans l'en-tête de l'onglet "
                f"« {onglet} ».\n"
                f"  Libellés acceptés : {', '.join(attendus)}\n"
                f"  En-tête réellement lu : {lus}\n"
                f"  L'ordre des colonnes n'a pas d'importance : seul le "
                f"libellé compte.\n"
                f"  Classeur NON modifié."
            )
        indices[role] = index
    return indices


def valeur(row, index):
    """Valeur de la colonne `index` dans `row`, ou None si elle est absente.

    `row` provient de `iter_rows(values_only=True)` : une ligne peut être
    plus courte que l'en-tête si les dernières cellules sont vides. On
    retourne None plutôt que de lever une IndexError.
    """
    if not row or index is None or index < 0 or index >= len(row):
        return None
    return row[index]


def aligner(valeurs, entete_source, entete_cible, onglet="?"):
    """Réordonne `valeurs` pour suivre l'ordre de `entete_cible`.

    Utilisé pour écrire dans un onglet d'exclusion unique des lignes
    provenant d'onglets de secteurs dont l'ordre de colonnes peut différer :
    sans ce réordonnancement, la première ligne rencontrée fixerait l'ordre
    et toutes les autres seraient stockées décalées.

    Lève une RuntimeError si les deux en-têtes ne portent pas exactement
    les mêmes colonnes — cela signifierait une perte de données, qui doit
    être signalée et non silencieuse.
    """
    cles_source = {}
    for i, libelle in enumerate(entete_source or ()):
        cle = _norm_libelle(libelle)
        if cle and cle not in cles_source:
            cles_source[cle] = i

    cible = []
    for libelle in entete_cible or ():
        cle = _norm_libelle(libelle)
        if cle and cle not in cible:
            cible.append(cle)

    manquantes = [c for c in cible if c not in cles_source]
    en_trop = [c for c in cles_source if c not in cible]
    if manquantes or en_trop:
        details = []
        if manquantes:
            details.append(
                "absentes de l'onglet source : " + ", ".join(manquantes))
        if en_trop:
            details.append(
                "absentes de l'en-tête de référence : " + ", ".join(en_trop))
        raise RuntimeError(
            f"En-têtes de colonnes incompatibles pour l'onglet « {onglet} » "
            f"({'; '.join(details)}).\n"
            f"  Les lignes ne peuvent pas être réordonnées sans perte de "
            f"données. Classeur NON modifié."
        )

    return [valeur(valeurs, cles_source[cle]) for cle in cible]