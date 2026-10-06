# Sommaire du projet — Analyse et rationalisation du parc applicatif (CNW)

> **Objectif de ce fichier** : Servir de référence de contexte lors de l'ouverture d'une nouvelle session. Il permet à l'agent de comprendre rapidement l'état du projet, les décisions prises et les artefacts disponibles.

---

## 1. Contexte organisationnel

| Élément | Détail |
|---|---|
| **Organisation** | Conseil de la Nation Huronne-Wendat (CNW) |
| **Pilier stratégique** | 7.4.4 — Harmonisation, rationalisation et optimisation de l'écosystème technologique (dette technologique) |
| **Répondant TI** | Donald Perron |

### Objectif du projet
Recenser, rationaliser et standardiser les logiciels de l'organisation afin que le département TI reprenne le contrôle sur le parc logiciel du CNW.

### Résultats attendus
- Cartographie complète des applications
- Vision claire des usages
- Base pour : décommissionnement, standardisation, documentation, sécurité et projets futurs

---

## 2. Phasage du projet

| Phase | Nom | Statut | Description |
|---|---|---|---|
| **Phase 0** | Gouvernance globale | Terminé | Structuration méthodologique, définition du périmètre, rôles et gabarits |
| **Phase 1** | Recensement structuré | En cours | Collecte des données, application des règles d'exclusion, préparation des entrevues |
| **Phase 2** | Analyse et consolidation | À venir | Analyse transversale, identification des doublons, shadow IT, outils critiques |
| **Phase 3** | Consolidation globale | À venir | Feuille de route, gouvernance continue |

---

## 3. Travaux d'exclusion

### 3.1 Fichier source principal
**`User_Workstation_App_Associations_Consolidated.xlsx`**
- Inventaire consolidé des associations utilisateurs / postes de travail / applications
- La version actuelle du fichier contient des onglets par **secteur** avec "TI" comme dernier secteur. Cette distinction est importante pour l'application des règles.

### 3.2 Règles d'exclusion appliquées

Six règles d'exclusion dans 6 fichiers (.MD) débutant par "Règle_" ont été définies pour nettoyer l'inventaire et ne conserver que les applications "d'affaires" pertinentes :

| Règle | Fichier | Description |
|---|---|---|---|
| **Règle 1** | `Regle_01_exclusions_generales.md` | Exclure tout ce qui est installé par défaut avec Windows 11 incluant les mises à niveau |
| **Règle 2** | `Regle_02_microsoft_standards.md` | Exclure les applications Microsoft standards (Office 365, Edge, Teams, etc.). **Exceptions** : Visio, Power BI, VS Code |
| **Règle 3** | `Regle_03_outils_installes_par_defaut.md` | Exclure les outils déployés par défaut par les TI (7-Zip, FortiClient, VLC, Adobe Reader, Antidote, etc.) |
| **Règle 4** | `Regle_04_pilotes_utilitaires_materiels.md` | Exclure pilotes, utilitaires matériels et composantes système (publisher commençant par "CN=", non-Cloud) |
| **Règle 5** | `Regle_05_dedoublonnage_sources.md` | Dédoublonner entre sources (priorité : Workstation > Cloud > Server) par UserName + WorkstationName + ApplicationName |
| **Règle 6** | `Regle_06_exclusions_specifiques.md` | Exclusions spécifiques par nom exact (1Password, Bonjour, KnowBe4, Lets Encrypt, Logi Plugin Service, etc.) |

---

## 4. Méthodologie de traitement Excel

### Principes directeurs
1. **Aucune perte de données** : Toute ligne supprimée d'un onglet secteur doit être archivée dans l'onglet **"Excluded Regles..."** avec la règle applicable
2. **Traitement séquentiel** : Les règles sont appliquées manuellement une à une, avec validation après chaque étape
3. **Gestion des ressources** : Sauvegarder le classeur après chaque règle pour libérer la mémoire Excel

---

*Dernière mise à jour : 2026-09-04*
